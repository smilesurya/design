"""
mcid_repair.py

Scans every page of a PDF for duplicate MCIDs (Marked Content IDs) and
automatically repairs every one it finds -- no page number or MCID value
needs to be supplied manually.

Background: PAC / Acrobat's accessibility checker require MCIDs to be
unique within a page's content stream. If two BDC marked-content sequences
share one MCID, PAC aborts with "MCID <n> already present." without saying
where. This script finds every occurrence on every page, keeps the first
occurrence of each duplicated MCID as-is, renumbers every repeat to a fresh
unused MCID on that page, and patches the matching structure-tree
references so the tags still point at the right content.

The original file is never modified -- output goes to <input>_fixed.pdf.

Usage:
    python mcid_repair.py input.pdf
"""

import sys
from collections import defaultdict

import pikepdf
from pikepdf import Name


def resolve_properties(operand, resources):
    """BDC's second operand is either an inline dict or a /Properties name
    referencing an entry in the page's resource dictionary."""
    if isinstance(operand, pikepdf.Dictionary):
        return operand

    if isinstance(operand, pikepdf.Name):
        props = resources.get("/Properties")

        if props and operand in props:
            return props[operand]

    return None


def extract_text_preview(instructions, start_index, max_chars=50):
    text = ""

    for operands, operator in instructions[start_index + 1:start_index + 60]:
        op = str(operator)

        if op == "EMC":
            break

        if op == "Tj" and operands:
            text += str(operands[0])

        elif op == "TJ" and operands:
            for el in operands[0]:
                if isinstance(el, pikepdf.String):
                    text += str(el)

        if len(text) >= max_chars:
            break

    return text[:max_chars].strip()


def scan_page(instructions, resources):
    """Return {mcid: [op_index, op_index, ...]} for every MCID on the page."""
    positions = defaultdict(list)

    for op_index, (operands, operator) in enumerate(instructions):
        if str(operator) != "BDC" or len(operands) < 2:
            continue

        props = resolve_properties(operands[1], resources)

        if props is None or "/MCID" not in props:
            continue

        positions[int(props["/MCID"])].append(op_index)

    return positions


def rewrite_content_stream(pdf, page, instructions, resources, fix_map):
    """fix_map: {op_index: new_mcid}"""
    new_instructions = []

    for op_index, (operands, operator) in enumerate(instructions):
        if op_index in fix_map:
            new_mcid = fix_map[op_index]
            operand1 = operands[1]

            if isinstance(operand1, pikepdf.Name):
                props = resolve_properties(operand1, resources)
                new_props = pikepdf.Dictionary(props)
                new_props["/MCID"] = new_mcid
                operands = pikepdf.Array([operands[0], new_props])

            else:
                operand1["/MCID"] = new_mcid
                operands = pikepdf.Array([operands[0], operand1])

        new_instructions.append((operands, operator))

    new_data = pikepdf.unparse_content_stream(new_instructions)
    page.Contents = pdf.make_stream(new_data)


def walk_struct_tree(node, page_ref, target_mcid, inherited_pg, matches, path):
    """DFS collecting every struct-tree reference to (page_ref, target_mcid),
    in document order, tracking inherited /Pg."""
    current_pg = (
        node.get("/Pg", inherited_pg)
        if isinstance(node, pikepdf.Dictionary)
        else inherited_pg
    )

    kids = node.get("/K") if isinstance(node, pikepdf.Dictionary) else None

    if kids is None:
        return

    is_array = isinstance(kids, pikepdf.Array)
    kid_list = kids if is_array else [kids]

    for i, kid in enumerate(kid_list):
        kid_path = path + [i]

        if isinstance(kid, int):
            if (
                current_pg is not None
                and current_pg == page_ref
                and kid == target_mcid
            ):
                matches.append(
                    {
                        "kind": "int",
                        "node": node,
                        "index": i,
                        "is_array": is_array,
                        "path": kid_path,
                    }
                )

        elif isinstance(kid, pikepdf.Dictionary):
            kid_type = kid.get("/Type")

            if kid_type == Name("/MCR"):
                pg = kid.get("/Pg", current_pg)

                if (
                    pg == page_ref
                    and int(kid.get("/MCID", -1)) == target_mcid
                ):
                    matches.append(
                        {
                            "kind": "mcr",
                            "node": kid,
                            "index": None,
                            "is_array": None,
                            "path": kid_path,
                        }
                    )

            elif kid_type == Name("/OBJR"):
                continue

            else:
                walk_struct_tree(
                    kid,
                    page_ref,
                    target_mcid,
                    current_pg,
                    matches,
                    kid_path,
                )


def apply_struct_fix(match, new_mcid):
    if match["kind"] == "int":
        node = match["node"]
        index = match["index"]

        if match["is_array"]:
            kid_list = node["/K"]
            kid_list[index] = new_mcid
            node["/K"] = kid_list

        else:
            node["/K"] = new_mcid

    else:
        match["node"]["/MCID"] = new_mcid


def process_pdf(pdf_path):
    pdf = pikepdf.open(pdf_path)
    struct_root = pdf.Root.get("/StructTreeRoot")
    total_fixed = 0
    pages_affected = 0

    for page_index, page in enumerate(pdf.pages):
        resources = page.get("/Resources", pikepdf.Dictionary({}))

        try:
            instructions = pikepdf.parse_content_stream(page)
        except Exception as e:
            print(
                f"⚠ Page {page_index + 1}: could not parse content stream: {e}"
            )
            continue

        positions = scan_page(instructions, resources)

        duplicated = {
            mcid: ops
            for mcid, ops in positions.items()
            if len(ops) > 1
        }

        if not duplicated:
            continue

        pages_affected += 1

        print(
            f"📄 Page {page_index + 1}: "
            f"{len(duplicated)} duplicated MCID value(s) found"
        )

        next_free_mcid = max(positions.keys()) + 1 if positions else 0
        fix_map = {}
        new_mcids_by_old = defaultdict(list)

        for mcid, op_indices in duplicated.items():
            for i, op_index in enumerate(op_indices):
                preview = extract_text_preview(instructions, op_index)
                marker = "keep" if i == 0 else "->"

                print(
                    f'   MCID {mcid} occurrence {i + 1} '
                    f'(op #{op_index}) [{marker}]: "{preview}..."'
                )

                if i == 0:
                    continue

                fix_map[op_index] = next_free_mcid
                new_mcids_by_old[mcid].append(next_free_mcid)
                total_fixed += 1
                next_free_mcid += 1

        rewrite_content_stream(
            pdf,
            page,
            instructions,
            resources,
            fix_map,
        )

        if struct_root is None:
            print(
                "   ⚠ No /StructTreeRoot -- content stream fixed, "
                "no tag tree to patch."
            )
            continue

        page_ref = page.obj

        for mcid, assigned_new_mcids in new_mcids_by_old.items():
            matches = []

            walk_struct_tree(
                struct_root,
                page_ref,
                mcid,
                None,
                matches,
                [],
            )

            expected = len(duplicated[mcid])

            if len(matches) != expected:
                print(
                    f"   ⚠ MCID {mcid}: expected {expected} "
                    f"struct-tree reference(s), found {len(matches)}. "
                    f"Tag tree NOT patched for this MCID -- check manually."
                )
                continue

            for match, new_mcid in zip(matches[1:], assigned_new_mcids):
                apply_struct_fix(match, new_mcid)

                print(
                    f"   ✅ Patched struct-tree reference at path "
                    f"{match['path']} -> MCID {new_mcid}"
                )

    if total_fixed == 0:
        print(
            "✅ No duplicate MCIDs found anywhere in the document. "
            "Nothing to fix."
        )
        pdf.close()
        return

    output_path = pdf_path.rsplit(".", 1)[0] + "_fixed.pdf"
    pdf.save(output_path)

    print(
        f"\n🚀 Fixed {total_fixed} duplicate MCID instance(s) "
        f"across {pages_affected} page(s)."
    )
    print(f"✅ Saved: {output_path}")

    pdf.close()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python mcid_repair.py input.pdf")
        sys.exit(1)

    process_pdf(sys.argv[1])
