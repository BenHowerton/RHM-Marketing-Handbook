#!/usr/bin/env python3
"""
Transform COBE Marketing Playbook Final Master.pptx based on Kate Watson's revisions.
Outputs COBE_Marketing_Playbook_KW_Revisions.pptx

Since python-pptx is not available, we work directly with the PPTX ZIP/XML structure.
"""

import zipfile
import xml.etree.ElementTree as ET
import re
import os
import shutil
import copy
from datetime import date

# Paths
INPUT_PATH = "/home/user/RHM-Marketing-Handbook/COBE Marketing Playbook Final Master.pptx"
OUTPUT_PATH = "/home/user/RHM-Marketing-Handbook/COBE_Marketing_Playbook_KW_Revisions.pptx"
WORK_DIR = "/tmp/pptx_work"

# Namespaces used in PPTX XML
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}

# Register all namespaces to preserve them during serialization
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')


def extract_pptx(pptx_path, work_dir):
    """Extract PPTX to working directory."""
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(pptx_path, 'r') as z:
        z.extractall(work_dir)


def repackage_pptx(work_dir, output_path):
    """Re-package working directory into PPTX."""
    if os.path.exists(output_path):
        os.remove(output_path)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk(work_dir):
            for f in files:
                file_path = os.path.join(root, f)
                arcname = os.path.relpath(file_path, work_dir)
                zout.write(file_path, arcname)


def read_xml(filepath):
    """Read and parse an XML file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    return content


def write_xml(filepath, content):
    """Write XML content to file."""
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)


def text_replace_in_xml(xml_content, old_text, new_text):
    """Replace text within <a:t> elements in XML string."""
    # This handles text within XML tags properly
    def replacer(match):
        original = match.group(0)
        return original.replace(old_text, new_text)

    # Match <a:t>...</a:t> elements and replace within them
    pattern = r'<a:t>[^<]*' + re.escape(old_text) + r'[^<]*</a:t>'
    result = re.sub(pattern, replacer, xml_content)
    return result


def global_text_replace(xml_content):
    """Apply all global find-and-replace operations."""

    # 1. COBECAST → Bronco Business Podcast (all cases)
    # First handle "COBECAST Podcast" → "Bronco Business Podcast" (avoid double "Podcast")
    xml_content = text_replace_in_xml(xml_content, 'COBECAST Podcast', 'Bronco Business Podcast')
    xml_content = text_replace_in_xml(xml_content, 'COBECAST podcast', 'Bronco Business Podcast')
    # Then handle remaining standalone COBECAST
    xml_content = text_replace_in_xml(xml_content, 'COBECAST', 'Bronco Business Podcast')
    xml_content = text_replace_in_xml(xml_content, 'Cobecast', 'Bronco Business Podcast')
    xml_content = text_replace_in_xml(xml_content, 'cobecast', 'Bronco Business Podcast')

    # 2. Localist → Boise State Events Calendar (Localist)
    # But "Localist Event Calendar" should become "Boise State Events Calendar"
    xml_content = text_replace_in_xml(xml_content, 'Localist Event Calendar', 'Boise State Events Calendar')
    xml_content = text_replace_in_xml(xml_content, 'Localist listing', 'Boise State Events Calendar listing')
    xml_content = text_replace_in_xml(xml_content, 'Localist', 'Boise State Events Calendar (Localist)')

    # 3. LinkedIn handle: @boisestatecobe in LinkedIn context → full name
    # Only change "LinkedIn (@boisestatecobe)" - keep Instagram handle
    xml_content = text_replace_in_xml(xml_content,
        'LinkedIn (@boisestatecobe)',
        'LinkedIn (College of Business and Economics at Boise State)')

    return xml_content


# ============================================================
# SLIDE-SPECIFIC MODIFICATION FUNCTIONS
# ============================================================

def fix_slide1_subtitle(xml_content):
    """Fix broken line break in slide 1 subtitle.
    Merge the two paragraphs about 'programs' and 'to leverage' into one.
    """
    # The subtitle is split into two paragraphs:
    # <a:t>A step-by-step guide for College of Business and Economics programs</a:t>
    # (paragraph break)
    # <a:t>to leverage COBE Marketing channels effectively and consistently.</a:t>

    # Step 1: Remove the second paragraph's text first (match <a:t> starting with "to leverage")
    xml_content = re.sub(
        r'<a:t>to leverage COBE Marketing channels effectively and consistently\.</a:t>',
        '<a:t></a:t>',
        xml_content
    )

    # Step 2: Now expand the first paragraph to include the full text
    xml_content = xml_content.replace(
        '<a:t>A step-by-step guide for College of Business and Economics programs</a:t>',
        '<a:t>A step-by-step guide for College of Business and Economics programs to leverage COBE Marketing channels effectively and consistently.</a:t>'
    )

    return xml_content


def fix_slide6_audiences(xml_content):
    """Update slide 6 subtitle to emphasize current students as primary audience."""
    # The apostrophe in "you're" is a Unicode right single quotation mark U+2019
    old_text = 'Know who you\u2019re talking to before you create content.'
    new_text = 'Current students are your primary audience. Most COBE marketers are trying to get current students involved in events, speakers, or programs. The audiences below (prospective students, employers, partners) are secondary and mostly handled by specialized teams with their own systems.'
    xml_content = text_replace_in_xml(xml_content, old_text, new_text)
    return xml_content


def fix_slide7_pillars(xml_content):
    """Update slide 7 subtitle to clarify RHM-specific examples."""
    xml_content = text_replace_in_xml(xml_content,
        'Every piece of content should connect to at least one of these themes.',
        'These messaging pillars were developed for the RHM certificate campaigns. They are RHM-specific examples that other programs can adapt to their own messaging needs \xe2\x80\x94 they are not universal requirements.')
    return xml_content


def fix_slide8_funnel(xml_content):
    """Add Hands-On context to slide 8 subtitle."""
    xml_content = text_replace_in_xml(xml_content,
        'Match your goal to the right funnel stage and channel.',
        'This funnel framework is primarily for Hands-On teams building their own campaigns. Hands-Off users can skip this \xe2\x80\x94 just tell COBE Marketing your goal and they will recommend the right channels.')
    return xml_content


def fix_slide12_table(xml_content):
    """Remove specific rows and channel references from slide 12 table.
    Remove: 'Recruit to certificate program' row, 'Build employer awareness' row
    Remove: 'Class Visits' and 'Paid Ads' from channel recommendations
    """
    # Parse as XML tree for table manipulation
    root = ET.fromstring(xml_content)

    # Find table
    tbl = root.find('.//' + '{http://schemas.openxmlformats.org/drawingml/2006/main}tbl')
    if tbl is not None:
        rows_to_remove = []
        for tr in tbl.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr'):
            texts = []
            for t_elem in tr.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
                if t_elem.text:
                    texts.append(t_elem.text)
            combined = ' '.join(texts)

            # Mark rows for removal
            if 'Recruit to certificate program' in combined:
                rows_to_remove.append(tr)
            elif 'Build employer awareness' in combined:
                rows_to_remove.append(tr)

        for tr in rows_to_remove:
            tbl.remove(tr)

        # Remove "Class Visits" and "Paid Ads" from remaining channel text
        for t_elem in tbl.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
            if t_elem.text:
                # Remove Class Visits
                t_elem.text = re.sub(r'\s*\+?\s*Class Visits\s*\+?\s*', lambda m: ' + ' if '+' in m.group() and m.start() > 0 and m.end() < len(t_elem.text) else '', t_elem.text)
                t_elem.text = re.sub(r',?\s*Class Visits\s*,?', '', t_elem.text)
                # Remove Paid Ads
                t_elem.text = re.sub(r'\s*\+?\s*Paid Ads\s*\+?\s*', lambda m: ' + ' if '+' in m.group() and m.start() > 0 and m.end() < len(t_elem.text) else '', t_elem.text)
                t_elem.text = re.sub(r',?\s*Paid Ads\s*,?', '', t_elem.text)
                # Clean up double + signs or trailing +
                t_elem.text = re.sub(r'\s*\+\s*\+\s*', ' + ', t_elem.text)
                t_elem.text = re.sub(r'\s*\+\s*$', '', t_elem.text)
                t_elem.text = re.sub(r'^\s*\+\s*', '', t_elem.text)

    # Serialize back
    return ET.tostring(root, encoding='unicode', xml_declaration=True)


def fix_slide13_timeline(xml_content):
    """Remove '1 WEEK OUT', 'DAY OF', and '1 WEEK AFTER' sections from slide 13.
    The content is in text boxes (shapes), not a table.
    """
    root = ET.fromstring(xml_content)

    # Find all shapes
    sp_tag = '{http://schemas.openxmlformats.org/presentationml/2006/main}sp'
    cSld = root.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld')
    if cSld is None:
        return xml_content

    spTree = cSld.find('{http://schemas.openxmlformats.org/presentationml/2006/main}spTree')
    if spTree is None:
        return xml_content

    shapes_to_remove = []

    # Identify shapes that contain the text we want to remove
    removal_keywords = ['1 WEEK OUT', '1 WEEK AFTER', 'DAY OF',
                        'Final headcount check', 'logistics confirmed',
                        'Reminder posts, stories, email reminder blast',
                        'Run the event, coordinate check-in',
                        'Live stories, day-of posts, photography',
                        'Share feedback, thank-you notes',
                        'Post-event recap, metrics report, photo gallery']

    for sp in spTree.findall(sp_tag):
        texts = []
        for t_elem in sp.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
            if t_elem.text:
                texts.append(t_elem.text)
        combined = ' '.join(texts).strip()

        for kw in removal_keywords:
            if kw in combined:
                shapes_to_remove.append(sp)
                break

    for sp in shapes_to_remove:
        spTree.remove(sp)

    return ET.tostring(root, encoding='unicode', xml_declaration=True)


def fix_slide14_email(xml_content):
    """Simplify the sample email on slide 14.
    Remove: 'Channels requested', 'Target dates', 'Success metric'
    """
    # Remove the specific lines
    xml_content = text_replace_in_xml(xml_content,
        'Channels requested: Student Events Newsletter, Instagram post + stories, MBEB screens',
        '')
    xml_content = text_replace_in_xml(xml_content,
        'Target dates: Promotion to begin Oct 1; event Oct 23',
        '')
    xml_content = text_replace_in_xml(xml_content,
        'Success metric: RSVPs via Handshake',
        '')
    return xml_content


def consolidate_slides_15_18(slide15_content):
    """Rewrite slide 15 to be a consolidated overview of all 4 Hands-Off channels.
    The individual detail slides (16, 17, 18) will be deleted.
    """
    root = ET.fromstring(slide15_content)

    # Find all shapes and their text
    sp_tag = '{http://schemas.openxmlformats.org/presentationml/2006/main}sp'
    cSld = root.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld')
    spTree = cSld.find('{http://schemas.openxmlformats.org/presentationml/2006/main}spTree')

    # Strategy: Keep the slide structure but rewrite the text content
    # We'll map specific text replacements for the main content areas

    replacements = {
        'Hands-Off: Instagram': 'Hands-Off Channel Overview',
        'BEST FOR': 'INSTAGRAM',
        'Student engagement, event awareness, program visibility. 1,436 followers.': 'Best for: Student engagement, event awareness, program visibility\nYou provide: Event/program details, photos (min 1080px), registration link',
        'ASSET TYPES': 'LINKEDIN',
        'Square posts (1080x1080), Stories/Reels (1080x1920)': 'Best for: Employer partnerships, alumni highlights, professional storytelling\nYou provide: Story/achievement details, outcome data, partner quotes/logos',
        'CADENCE': 'EMAIL NEWSLETTER',
        '3-4 posts/week during peak recruitment; stories around events': 'Best for: Driving event attendance, applications, program inquiries (5,000+ subscribers)\nYou provide: Rectangle landscape photo, short paragraph with value props, CTA link',
        'LEAD TIME': 'INTERNAL CHANNELS',
        '10 business days; urgent posts need marketing manager approval': 'Best for: High-visibility to current COBE students (MBEB screens, classroom slides, campus screens)\nYou provide: Event graphic (1920x1080 JPG, dark bg), date range, key details',
        'WHAT YOU SEND COBE': '',
        'Event/program details, photos (min 1080px), registration link, goal/success metric.': 'For all channels: Email cobe-info@boisestate.edu with event details. Standard lead time is 10 business days.',
        'PRO TIPS': '',
        'Reels outperform static posts for reach. Use Link Stickers in Stories. Always include a clear CTA.': '',
        'FUNNEL STAGE': '',
        'Awareness + Consideration. Best for current students and prospective undergrads.': '',
    }

    for sp in spTree.findall(sp_tag):
        for t_elem in sp.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
            if t_elem.text and t_elem.text in replacements:
                t_elem.text = replacements[t_elem.text]

    return ET.tostring(root, encoding='unicode', xml_declaration=True)


def fix_slide19_podcast(xml_content):
    """Add note that Bronco Business Podcast is a special case."""
    # The global replace already handles COBECAST → Bronco Business Podcast
    # Add note about special case
    xml_content = text_replace_in_xml(xml_content,
        'Employer partnerships, alumni engagement, thought leadership, student recruiting influence.',
        'Employer partnerships, alumni engagement, thought leadership, student recruiting influence. NOTE: The Bronco Business Podcast is a special case that requires direct coordination with Kate Watson. It is not a standard marketing channel offered broadly.')
    return xml_content


def fix_slide20_podcast(xml_content):
    """Add note to slide 20 about special case."""
    # The period after "(not promotion)" may be in a separate XML element,
    # so match without the trailing period
    xml_content = text_replace_in_xml(xml_content,
        'Industry professional, alumni, employer, or community partner who offers learning, career insight, leadership, or industry perspective (not promotion)',
        'Industry professional, alumni, employer, or community partner who offers learning, career insight, leadership, or industry perspective (not promotion). Note: This process requires direct coordination with Kate Watson')
    return xml_content


def fix_slide25_calendar(xml_content):
    """Add note about coordinating with Kim and COBE social calendar."""
    # Add to the end of the how-to section
    xml_content = text_replace_in_xml(xml_content,
        'Submit completed calendar with as',
        'NOTE: Kim Presnell maintains her own COBE social schedule. Your submitted calendar will be coordinated with Kim and the main COBE social calendar to ensure alignment. Submit completed calendar with as')

    # If the above didn't match (truncated text), try alternate approach
    if 'NOTE: Kim Presnell maintains' not in xml_content:
        xml_content = text_replace_in_xml(xml_content,
            'Review against recommended cadence below',
            'Review against recommended cadence below\nNOTE: Kim Presnell maintains her own COBE social schedule. Your submitted calendar will be coordinated with Kim and the main COBE social calendar to ensure alignment.')

    return xml_content


def fix_slide29_photo(xml_content):
    """Change 'Square-ish photo' to 'Rectangle landscape photo'."""
    xml_content = text_replace_in_xml(xml_content,
        'Square-ish photo',
        'Rectangle landscape photo')
    return xml_content


def fix_slide31_wordpress(xml_content):
    """Update WordPress section for specialized teams only."""
    xml_content = text_replace_in_xml(xml_content,
        'WORDPRESS ACCESS REQUIREMENTS',
        'WORDPRESS ACCESS (Specialized Teams Only)')

    # Add note about submitting requests to cobe-info
    xml_content = text_replace_in_xml(xml_content,
        'Complete web accessibility compliance training (60-90 min,',
        'Most programs should submit web update requests to cobe-info@boisestate.edu rather than pursuing WordPress access directly. The university is reducing the number of WordPress admins and webpages. For specialized teams: Complete web accessibility compliance training (60-90 min,')

    return xml_content


def fix_slide50_tracking(xml_content):
    """Simplify tracking guidance: QR codes for everyone, UTMs/Looker for specialized teams."""
    xml_content = text_replace_in_xml(xml_content,
        'Track with Flowcode + UTMs + Looker Studio',
        'Track with QR codes (everyone); UTMs + Looker Studio (specialized teams)')

    xml_content = text_replace_in_xml(xml_content,
        'Flowcode QR per physical placement. UTM parameters on every digital link. Looker Studio dashboard compares channels side-by-side.',
        'For general programs: Create Flowcode QR codes for every physical placement (posters, flyers, handouts) \xe2\x80\x94 this is the simplest and most effective tracking method. For specialized teams: Also set up UTM parameters on digital links and use Looker Studio dashboards for side-by-side channel comparison.')

    return xml_content


def fix_slide56_contacts(xml_content):
    """Fix contacts table: remove McKenna, Ann, Ashley, Kent rows. Add Central Marketing."""
    root = ET.fromstring(xml_content)

    # Find table
    tbl = root.find('.//' + '{http://schemas.openxmlformats.org/drawingml/2006/main}tbl')
    if tbl is None:
        return xml_content

    rows_to_remove = []
    last_kept_row = None

    for tr in tbl.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr'):
        texts = []
        for t_elem in tr.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
            if t_elem.text:
                texts.append(t_elem.text)
        combined = ' '.join(texts)

        # Remove rows
        if 'McKenna Howard' in combined:
            rows_to_remove.append(tr)
        elif 'Ann Hottinger' in combined:
            rows_to_remove.append(tr)
        elif 'Ashley Walter' in combined:
            rows_to_remove.append(tr)
        elif 'Kent Neupert' in combined:
            rows_to_remove.append(tr)
        else:
            last_kept_row = tr

    for tr in rows_to_remove:
        tbl.remove(tr)

    # Add "Boise State Central Marketing" row
    # Clone the last kept row as a template for the new row
    if last_kept_row is not None:
        new_row = copy.deepcopy(last_kept_row)
        # Find all text cells and set their content
        cells = new_row.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tc')
        cell_texts = ['Design Support', 'Boise State Central Marketing',
                      'Graphic design requests (for programs not already working with COBE design staff)',
                      'boisestate.edu/communicationsandmarketing/creative/request-design-support/']

        for ci, cell in enumerate(cells):
            if ci < len(cell_texts):
                # Clear existing text and set new
                for t_elem in cell.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
                    t_elem.text = cell_texts[ci]
                    break  # Only set first text element per cell

        tbl.append(new_row)

    return ET.tostring(root, encoding='unicode', xml_declaration=True)


def fix_slide57_brand(xml_content):
    """Update brand reference to link to brand website."""
    # Add brand website reference
    xml_content = text_replace_in_xml(xml_content,
        'KEY GUIDELINES',
        'KEY GUIDELINES (Digital Colors Shown \xe2\x80\x94 See Full Palette at boisestate.edu/brand/visual-identity/)')

    xml_content = text_replace_in_xml(xml_content,
        'Typography: Garamond (body), Gotham/Oswald (headlines). Web: Oswald + Lato.',
        'For the complete color palette (including print), typography rules, and logo usage, visit: boisestate.edu/brand/visual-identity/')

    return xml_content


def fix_slide59_faq(xml_content):
    """Fix FAQ slide: indent answers, remove Marq refs, note Canva transition, fix Resource Library reference."""
    # Indent answers with arrows/dashes
    answer_replacements = {
        'No. All content goes through cobe-info for review.':
            '\xe2\x86\x92 No. All content goes through cobe-info for review.',
        'Standard is 10 business days. Rush = director approval.':
            '\xe2\x86\x92 Standard is 10 business days. Rush = director approval.',
        'Hands-Off: COBE creates them. Hands-On: use templates.':
            '\xe2\x86\x92 Hands-Off: COBE creates them. Hands-On: use brand templates in Canva (replacing Marq campus-wide).',
        'Yes! Email cobe-info, subject: STUDENT ORG REQUEST.':
            '\xe2\x86\x92 Yes! Email cobe-info, subject: STUDENT ORG REQUEST.',
        'See the Resource Library slide or the playbook repository.':
            '\xe2\x86\x92 Check with COBE Marketing at cobe-info@boisestate.edu for current templates. Canva will replace Marq as the campus-wide template platform.',
        'Contact Kate Watson directly for urgent changes or crisis situations.':
            '\xe2\x86\x92 Contact Kate Watson directly for urgent changes or crisis situations.',
    }

    for old, new in answer_replacements.items():
        xml_content = text_replace_in_xml(xml_content, old, new)

    return xml_content


def fix_slide61_changelog(xml_content):
    """Add v4.0 changelog entry."""
    root = ET.fromstring(xml_content)

    # Find table
    tbl = root.find('.//' + '{http://schemas.openxmlformats.org/drawingml/2006/main}tbl')
    if tbl is None:
        return xml_content

    rows = tbl.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tr')

    # Find the first empty row (row 5 based on our earlier inspection)
    today = date.today().strftime('%Y-%m-%d')
    new_entry = ['v4.0', today, 'Ben Howerton',
                 'Kate Watson revisions: consolidated hands-off channels, fixed audience focus, renamed Bronco Business Podcast, simplified for department chairs, updated contacts and brand references.']

    for tr in rows:
        texts = []
        for t_elem in tr.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
            if t_elem.text:
                texts.append(t_elem.text.strip())

        # Find first empty row
        if not any(texts):
            # Fill in this row
            cells = tr.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}tc')
            for ci, cell in enumerate(cells):
                if ci < len(new_entry):
                    for t_elem in cell.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
                        t_elem.text = new_entry[ci]
                        break
            break

    return ET.tostring(root, encoding='unicode', xml_declaration=True)


def delete_slide(work_dir, slide_num):
    """Delete a slide from the PPTX by removing its files and references.
    slide_num is the file number (e.g., 16 for slide16.xml).
    """
    # Remove slide file
    slide_path = os.path.join(work_dir, f'ppt/slides/slide{slide_num}.xml')
    if os.path.exists(slide_path):
        os.remove(slide_path)

    # Remove slide rels file
    rels_path = os.path.join(work_dir, f'ppt/slides/_rels/slide{slide_num}.xml.rels')
    if os.path.exists(rels_path):
        os.remove(rels_path)

    # Remove notes slide and its rels
    notes_path = os.path.join(work_dir, f'ppt/notesSlides/notesSlide{slide_num}.xml')
    if os.path.exists(notes_path):
        os.remove(notes_path)
    notes_rels_path = os.path.join(work_dir, f'ppt/notesSlides/_rels/notesSlide{slide_num}.xml.rels')
    if os.path.exists(notes_rels_path):
        os.remove(notes_rels_path)

    # Remove from presentation.xml
    pres_path = os.path.join(work_dir, 'ppt/presentation.xml')
    pres_content = read_xml(pres_path)

    # Find the rId for this slide from presentation.xml.rels
    pres_rels_path = os.path.join(work_dir, 'ppt/_rels/presentation.xml.rels')
    pres_rels_content = read_xml(pres_rels_path)

    # Find the rId that points to this slide
    rid_pattern = r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="slides/slide' + str(slide_num) + r'\.xml"[^>]*/>'
    rid_match = re.search(rid_pattern, pres_rels_content)

    if rid_match:
        rid = rid_match.group(1)

        # Remove the relationship entry
        pres_rels_content = re.sub(rid_pattern, '', pres_rels_content)
        write_xml(pres_rels_path, pres_rels_content)

        # Remove the sldId entry from presentation.xml
        # Pattern: <p:sldId id="NNN" r:id="rIdXX"/>
        sldid_pattern = r'<p:sldId[^>]*r:id="' + rid + r'"[^>]*/>'
        pres_content = re.sub(sldid_pattern, '', pres_content)
        write_xml(pres_path, pres_content)

    # Remove from [Content_Types].xml
    ct_path = os.path.join(work_dir, '[Content_Types].xml')
    ct_content = read_xml(ct_path)

    # Remove slide override
    ct_pattern = r'<Override[^>]*PartName="/ppt/slides/slide' + str(slide_num) + r'\.xml"[^>]*/>'
    ct_content = re.sub(ct_pattern, '', ct_content)

    # Remove notes slide override
    ct_notes_pattern = r'<Override[^>]*PartName="/ppt/notesSlides/notesSlide' + str(slide_num) + r'\.xml"[^>]*/>'
    ct_content = re.sub(ct_notes_pattern, '', ct_content)

    write_xml(ct_path, ct_content)


def update_slide_page_numbers(work_dir):
    """Update displayed page numbers on all remaining slides after deletions.
    Reads the presentation.xml to get current slide order, then updates
    the page number text on each slide.
    """
    pres_path = os.path.join(work_dir, 'ppt/presentation.xml')
    pres_content = read_xml(pres_path)
    pres_rels_path = os.path.join(work_dir, 'ppt/_rels/presentation.xml.rels')
    pres_rels_content = read_xml(pres_rels_path)

    # Build rId → slide file mapping
    rid_to_file = {}
    for m in re.finditer(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="(slides/slide\d+\.xml)"', pres_rels_content):
        rid_to_file[m.group(1)] = m.group(2)

    # Get slide order from presentation.xml
    slide_order = []
    for m in re.finditer(r'<p:sldId[^>]*r:id="(rId\d+)"', pres_content):
        rid = m.group(1)
        if rid in rid_to_file:
            slide_order.append(rid_to_file[rid])

    print(f"Updating page numbers for {len(slide_order)} slides")

    # Track which slides are section headers (no page numbers) vs content slides
    # Section headers typically don't have page number indicators
    # We'll update any slide that has a standalone number matching its old position

    # For each slide, check if it has a page number and update it
    content_page = 0  # Running page counter
    for position, slide_file in enumerate(slide_order, 1):
        slide_path = os.path.join(work_dir, 'ppt', slide_file)
        if not os.path.exists(slide_path):
            continue

        xml_content = read_xml(slide_path)

        # Check if this slide has text content (vs being a section header)
        all_texts = re.findall(r'<a:t>([^<]*)</a:t>', xml_content)
        non_empty = [t for t in all_texts if t.strip()]

        # Section headers and title slides typically have fewer text elements
        # and don't have a standalone page number
        # Look for a shape that contains only a number (the page number indicator)

        # For now, just track position
        content_page += 1

    # Actually, the page numbers in this deck follow a pattern where:
    # - Title slide (pos 1): no number
    # - Section headers: no number
    # - Content slides: sequential page number matching their position
    #
    # After deleting slides, we need the remaining slides to have
    # sequential numbers. The simplest approach: update page numbers
    # to match the new position in the deck.

    for position, slide_file in enumerate(slide_order, 1):
        slide_path = os.path.join(work_dir, 'ppt', slide_file)
        if not os.path.exists(slide_path):
            continue

        xml_content = read_xml(slide_path)

        # Find shapes that contain only a page number (1-3 digit number as sole content)
        # These are typically small shapes with just the number
        shape_pattern = r'(<p:sp\b[^>]*>.*?</p:sp>)'

        def update_page_number_shape(match):
            shape_xml = match.group(0)
            texts = re.findall(r'<a:t>([^<]*)</a:t>', shape_xml)
            non_empty = [t.strip() for t in texts if t.strip()]

            # If this shape contains exactly one text element that's a 1-3 digit number
            if len(non_empty) == 1 and re.match(r'^\d{1,3}$', non_empty[0]):
                old_num = non_empty[0]
                new_num = str(position)
                # Replace the number
                shape_xml = re.sub(
                    r'(<a:t>)\s*' + old_num + r'\s*(</a:t>)',
                    r'\g<1>' + new_num + r'\g<2>',
                    shape_xml
                )
            return shape_xml

        xml_content = re.sub(shape_pattern, update_page_number_shape, xml_content, flags=re.DOTALL)
        write_xml(slide_path, xml_content)


def verify_no_bad_terms(work_dir):
    """Final verification pass: check for remaining COBECAST, bare Localist, @boisestatecobe on LinkedIn."""
    issues = []

    pres_rels_path = os.path.join(work_dir, 'ppt/_rels/presentation.xml.rels')
    pres_rels_content = read_xml(pres_rels_path)
    pres_path = os.path.join(work_dir, 'ppt/presentation.xml')
    pres_content = read_xml(pres_path)

    # Get list of remaining slides
    remaining_slides = []
    for m in re.finditer(r'Target="(slides/slide\d+\.xml)"', pres_rels_content):
        remaining_slides.append(m.group(1))

    for slide_file in sorted(remaining_slides):
        slide_path = os.path.join(work_dir, 'ppt', slide_file)
        if not os.path.exists(slide_path):
            continue

        xml_content = read_xml(slide_path)
        texts = re.findall(r'<a:t>([^<]*)</a:t>', xml_content)

        for t in texts:
            if 'COBECAST' in t or 'Cobecast' in t or 'cobecast' in t:
                issues.append(f"  {slide_file}: COBECAST found: '{t}'")

            # Check for bare "Localist" (not preceded by "Calendar (")
            if 'Localist' in t and 'Events Calendar' not in t and 'Calendar (Localist)' not in t:
                issues.append(f"  {slide_file}: Bare Localist found: '{t}'")

            # Check for @boisestatecobe in LinkedIn context
            if 'LinkedIn' in t and '@boisestatecobe' in t:
                issues.append(f"  {slide_file}: LinkedIn @boisestatecobe found: '{t}'")

    if issues:
        print("VERIFICATION ISSUES FOUND:")
        for i in issues:
            print(i)
        return False
    else:
        print("VERIFICATION PASSED: No remaining COBECAST, bare Localist, or LinkedIn @boisestatecobe")
        return True


# ============================================================
# MAIN EXECUTION
# ============================================================

def main():
    print("=" * 60)
    print("COBE Marketing Playbook Transformation Script")
    print("=" * 60)

    # Step 1: Extract PPTX
    print("\n[1/8] Extracting PPTX...")
    extract_pptx(INPUT_PATH, WORK_DIR)

    # Step 2: Apply global replacements to ALL slide XMLs
    print("\n[2/8] Applying global text replacements...")
    slides_dir = os.path.join(WORK_DIR, 'ppt/slides')
    for f in sorted(os.listdir(slides_dir)):
        if f.endswith('.xml') and f.startswith('slide'):
            fpath = os.path.join(slides_dir, f)
            content = read_xml(fpath)
            content = global_text_replace(content)
            write_xml(fpath, content)

    # Also apply to notes slides
    notes_dir = os.path.join(WORK_DIR, 'ppt/notesSlides')
    if os.path.exists(notes_dir):
        for f in sorted(os.listdir(notes_dir)):
            if f.endswith('.xml') and f.startswith('notesSlide'):
                fpath = os.path.join(notes_dir, f)
                content = read_xml(fpath)
                content = global_text_replace(content)
                write_xml(fpath, content)

    print("  Global replacements applied to all slides and notes.")

    # Step 3: Apply slide-specific changes
    print("\n[3/8] Applying slide-specific changes...")

    # Slide 1 - Fix subtitle line break
    print("  Slide 1: Fixing subtitle line break...")
    s1_path = os.path.join(slides_dir, 'slide1.xml')
    content = read_xml(s1_path)
    content = fix_slide1_subtitle(content)
    write_xml(s1_path, content)

    # Slide 6 - Target Audiences
    print("  Slide 6: Updating audience focus...")
    s6_path = os.path.join(slides_dir, 'slide6.xml')
    content = read_xml(s6_path)
    content = fix_slide6_audiences(content)
    write_xml(s6_path, content)

    # Slide 7 - Core Messaging Pillars
    print("  Slide 7: Clarifying RHM-specific pillars...")
    s7_path = os.path.join(slides_dir, 'slide7.xml')
    content = read_xml(s7_path)
    content = fix_slide7_pillars(content)
    write_xml(s7_path, content)

    # Slide 8 - Marketing Funnel
    print("  Slide 8: Adding Hands-On context...")
    s8_path = os.path.join(slides_dir, 'slide8.xml')
    content = read_xml(s8_path)
    content = fix_slide8_funnel(content)
    write_xml(s8_path, content)

    # Slide 12 - Channel table
    print("  Slide 12: Cleaning up channel table...")
    s12_path = os.path.join(slides_dir, 'slide12.xml')
    content = read_xml(s12_path)
    content = fix_slide12_table(content)
    write_xml(s12_path, content)

    # Slide 13 - Timeline
    print("  Slide 13: Removing late-timeline rows...")
    s13_path = os.path.join(slides_dir, 'slide13.xml')
    content = read_xml(s13_path)
    content = fix_slide13_timeline(content)
    write_xml(s13_path, content)

    # Slide 14 - Sample email
    print("  Slide 14: Simplifying request email...")
    s14_path = os.path.join(slides_dir, 'slide14.xml')
    content = read_xml(s14_path)
    content = fix_slide14_email(content)
    write_xml(s14_path, content)

    # Slides 15-18 - Consolidate
    print("  Slides 15-18: Consolidating into one overview...")
    s15_path = os.path.join(slides_dir, 'slide15.xml')
    content = read_xml(s15_path)
    content = consolidate_slides_15_18(content)
    write_xml(s15_path, content)

    # Slides 19-20 - Podcast notes
    print("  Slides 19-20: Adding podcast special case note...")
    s19_path = os.path.join(slides_dir, 'slide19.xml')
    content = read_xml(s19_path)
    content = fix_slide19_podcast(content)
    write_xml(s19_path, content)

    s20_path = os.path.join(slides_dir, 'slide20.xml')
    content = read_xml(s20_path)
    content = fix_slide20_podcast(content)
    write_xml(s20_path, content)

    # Slide 25 - Social calendar note
    print("  Slide 25: Adding Kim coordination note...")
    s25_path = os.path.join(slides_dir, 'slide25.xml')
    content = read_xml(s25_path)
    content = fix_slide25_calendar(content)
    write_xml(s25_path, content)

    # Slide 29 - Photo type
    print("  Slide 29: Fixing photo type...")
    s29_path = os.path.join(slides_dir, 'slide29.xml')
    content = read_xml(s29_path)
    content = fix_slide29_photo(content)
    write_xml(s29_path, content)

    # Slide 31 - WordPress
    print("  Slide 31: Updating WordPress section...")
    s31_path = os.path.join(slides_dir, 'slide31.xml')
    content = read_xml(s31_path)
    content = fix_slide31_wordpress(content)
    write_xml(s31_path, content)

    # Slide 50 - Tracking
    print("  Slide 50: Simplifying tracking guidance...")
    s50_path = os.path.join(slides_dir, 'slide50.xml')
    content = read_xml(s50_path)
    content = fix_slide50_tracking(content)
    write_xml(s50_path, content)

    # Slide 56 (displayed as 54) - Contacts
    print("  Slide 56 (Contacts): Updating contacts table...")
    s56_path = os.path.join(slides_dir, 'slide56.xml')
    content = read_xml(s56_path)
    content = fix_slide56_contacts(content)
    write_xml(s56_path, content)

    # Slide 57 (displayed as 55) - Brand
    print("  Slide 57 (Brand): Adding brand website reference...")
    s57_path = os.path.join(slides_dir, 'slide57.xml')
    content = read_xml(s57_path)
    content = fix_slide57_brand(content)
    write_xml(s57_path, content)

    # Slide 59 (displayed as 57) - FAQ
    print("  Slide 59 (FAQ): Fixing answers and references...")
    s59_path = os.path.join(slides_dir, 'slide59.xml')
    content = read_xml(s59_path)
    content = fix_slide59_faq(content)
    write_xml(s59_path, content)

    # Slide 61 - Changelog
    print("  Slide 61 (Changelog): Adding v4.0 entry...")
    s61_path = os.path.join(slides_dir, 'slide61.xml')
    content = read_xml(s61_path)
    content = fix_slide61_changelog(content)
    write_xml(s61_path, content)

    # Step 4: Delete slides (16, 17, 18, 33)
    print("\n[4/8] Deleting slides 16, 17, 18, 33...")
    for slide_num in [16, 17, 18, 33]:
        print(f"  Deleting slide {slide_num}...")
        delete_slide(WORK_DIR, slide_num)

    # Step 5: Update page numbers
    print("\n[5/8] Updating slide page numbers...")
    update_slide_page_numbers(WORK_DIR)

    # Step 6: Verification
    print("\n[6/8] Running verification...")
    verify_no_bad_terms(WORK_DIR)

    # Step 7: Repackage
    print("\n[7/8] Repackaging PPTX...")
    repackage_pptx(WORK_DIR, OUTPUT_PATH)

    # Step 8: Verify output
    print("\n[8/8] Verifying output file...")
    if os.path.exists(OUTPUT_PATH):
        size = os.path.getsize(OUTPUT_PATH)
        print(f"  Output: {OUTPUT_PATH}")
        print(f"  Size: {size:,} bytes")

        # Count slides in output
        with zipfile.ZipFile(OUTPUT_PATH, 'r') as z:
            slide_files = [n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
            print(f"  Slides: {len(slide_files)} (was 62, deleted 4)")
    else:
        print("  ERROR: Output file not created!")

    # Cleanup
    shutil.rmtree(WORK_DIR)

    print("\n" + "=" * 60)
    print("TRANSFORMATION COMPLETE")
    print("=" * 60)


if __name__ == '__main__':
    main()
