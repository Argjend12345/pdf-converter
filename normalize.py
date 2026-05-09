import zipfile
import shutil
import os
import sys
import traceback
import xml.etree.ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


START_MARKER = "Terms and Conditions"
STOP_MARKER = "NOTICE OF CANCELLATION"


def paragraph_text(p):
    """
    Combine all text runs in a paragraph.
    """
    return ''.join(
        t.text or ''
        for t in p.findall('.//w:t', NS)
    ).strip()


def paragraph_has_bold_text(p, target):
    """
    Detect bolded text even if split across runs.
    """

    full_text = paragraph_text(p)

    if target not in full_text:
        return False

    # Ensure at least one run is bold
    for r in p.findall('w:r', NS):
        rPr = r.find('w:rPr', NS)

        if rPr is not None and rPr.find('w:b', NS) is not None:
            return True

    return False


def normalize(path):
    print("Starting normalization...")
    print("Input file:", path)

    tmp = path + ".bak"

    print("Creating backup...")
    shutil.copy(path, tmp)

    print("Opening DOCX zip...")
    with zipfile.ZipFile(tmp, 'r') as z:
        names = z.namelist()
        files = {n: z.read(n) for n in names}

    if 'word/document.xml' not in files:
        raise Exception("word/document.xml not found in DOCX")

    print("Loading XML...")
    xml_str = files['word/document.xml'].decode('utf-8')

    ET.register_namespace('w', W_NS)

    root = ET.fromstring(xml_str)

    print("Finding paragraphs...")
    paragraphs = root.findall('.//w:p', NS)

    print("Paragraph count:", len(paragraphs))

    apply_changes = False
    modified_count = 0

    print(f"Searching for start marker: '{START_MARKER}'")
    print(f"Searching for stop marker: '{STOP_MARKER}'")

    for i, p in enumerate(paragraphs):

        text = paragraph_text(p)

        # START applying formatting
        if not apply_changes:
            if paragraph_has_bold_text(p, START_MARKER):
                apply_changes = True
                print(f"Found START marker at paragraph {i}")
                continue

        # Skip until start marker found
        if not apply_changes:
            continue

        # STOP applying formatting
        if paragraph_has_bold_text(p, STOP_MARKER):
            print(f"Found STOP marker at paragraph {i}")
            break

        # Find or create paragraph properties
        pPr = p.find('w:pPr', NS)

        if pPr is None:
            pPr = ET.SubElement(
                p,
                f'{{{W_NS}}}pPr'
            )

        # Find or create spacing element
        spacing = pPr.find('w:spacing', NS)

        if spacing is None:
            spacing = ET.SubElement(
                pPr,
                f'{{{W_NS}}}spacing'
            )

        # Set line spacing
        spacing.set(f'{{{W_NS}}}line', '253')
        spacing.set(f'{{{W_NS}}}lineRule', 'exact')

        # Optional:
        # spacing.set(f'{{{W_NS}}}after', '0')

        modified_count += 1

    print(f"Modified {modified_count} paragraphs")

    print("Writing updated XML...")
    new_xml = ET.tostring(
        root,
        encoding='utf-8',
        xml_declaration=True
    )

    files['word/document.xml'] = new_xml

    print("Rebuilding DOCX...")
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name in names:
            z.writestr(name, files[name])

    print("Removing temp backup...")
    os.remove(tmp)

    print("Normalization complete.")


if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            normalize(sys.argv[1])
        else:
            print("Usage: python normalize.py yourfile.docx")

    except Exception:
        print("ERROR DURING NORMALIZATION:")
        traceback.print_exc()
        sys.exit(1)