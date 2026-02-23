#!/usr/bin/env python3
"""
Apply Kate Watson's revisions to the COBE Marketing Playbook.
Works directly with PPTX ZIP/XML structure using only standard library.
"""

import zipfile
import re
import os

INPUT_PATH = "/home/user/RHM-Marketing-Handbook/COBE Marketing Playbook Final Master.pptx"
OUTPUT_PATH = "/home/user/RHM-Marketing-Handbook/COBE_Marketing_Playbook_KW_Revisions.pptx"

# Slides to delete (file numbers): 16, 17, 18, 33
SLIDES_TO_DELETE = {16, 17, 18, 33}

# rIds for deleted slides (from presentation.xml.rels)
DELETED_RIDS = {'rId21', 'rId22', 'rId23', 'rId38'}

# Files to skip entirely (deleted slides + their notes + rels)
FILES_TO_SKIP = set()
for s in SLIDES_TO_DELETE:
    FILES_TO_SKIP.add(f'ppt/slides/slide{s}.xml')
    FILES_TO_SKIP.add(f'ppt/slides/_rels/slide{s}.xml.rels')
    FILES_TO_SKIP.add(f'ppt/notesSlides/notesSlide{s}.xml')
    FILES_TO_SKIP.add(f'ppt/notesSlides/_rels/notesSlide{s}.xml.rels')


def global_text_replacements(content, filename=""):
    """Apply all global find-and-replace operations to XML content."""

    # 1. COBECAST → Bronco Business Podcast (case-sensitive variants)
    # Handle "COBECAST Podcast" → "Bronco Business Podcast" (not "Bronco Business Podcast Podcast")
    content = content.replace('COBECAST Podcast', 'Bronco Business Podcast')
    content = content.replace('COBECAST podcast', 'Bronco Business Podcast')
    content = content.replace('COBECAST', 'Bronco Business Podcast')
    content = content.replace('Cobecast', 'Bronco Business Podcast')
    content = content.replace('cobecast', 'Bronco Business Podcast')
    # Clean up any remaining double "Podcast Podcast"
    content = content.replace('Podcast Podcast', 'Podcast')
    content = content.replace('podcast Podcast', 'Podcast')

    # 2. Localist → Boise State Events Calendar
    # Handle "Localist Event Calendar" first (would become redundant otherwise)
    content = content.replace('Localist Event Calendar', 'Boise State Events Calendar')
    content = content.replace('Localist event calendar', 'Boise State Events Calendar')
    # Handle "Localist listing" specifically
    content = content.replace('Localist listing', 'Boise State Events Calendar listing')
    # Handle all remaining bare "Localist"
    content = content.replace('Localist', 'Boise State Events Calendar (Localist)')
    # Fix double-replacement chains
    content = content.replace('Boise State Events Calendar (Localist) Event Calendar', 'Boise State Events Calendar')
    content = content.replace('Boise State Events Calendar (Boise State Events Calendar (Localist))', 'Boise State Events Calendar')
    # Fix "Boise State Events Calendar (Localist) listing" leaving nested parens
    content = content.replace('Boise State Events Calendar (Localist) listing', 'Boise State Events Calendar listing')

    # 3. LinkedIn handle fix
    content = content.replace(
        'LinkedIn (@boisestatecobe)',
        'LinkedIn (College of Business and Economics at Boise State)'
    )

    return content


def fix_slide1(content):
    """Fix broken line break in subtitle."""
    # The subtitle has "programs" in one run, then a break, then "to leverage" in next run
    # Try multiple patterns for how the break might appear in XML

    # Pattern 1: <a:br/> between runs
    content = re.sub(
        r'(programs)</a:t></a:r><a:br/><a:r>([^<]*<a:rPr[^>]*/>[^<]*)<a:t>(to leverage)',
        r'\1 \3</a:t></a:r><a:r>\2<a:t>',
        content, flags=re.DOTALL
    )

    # Pattern 2: <a:br> with rPr child
    content = re.sub(
        r'(programs)</a:t></a:r><a:br><a:rPr[^/]*/></a:br><a:r>([^<]*)<a:t>(to leverage)',
        r'\1 \3</a:t></a:r><a:r>\2<a:t>',
        content, flags=re.DOTALL
    )

    # Pattern 3: separate paragraphs
    content = re.sub(
        r'(programs)</a:t>(</a:r>)(</a:p><a:p>)(<a:pPr[^>]*>(?:(?!</a:pPr>).)*?</a:pPr>)?(<a:r>)((?:(?!</a:r>).)*?)<a:t>(to leverage)',
        r'\1 \7</a:t>\2\3\4\5\6<a:t>',
        content, flags=re.DOTALL
    )

    # Pattern 4: endParaRPr then new paragraph
    content = re.sub(
        r'(programs)</a:t></a:r><a:endParaRPr[^/]*/></a:p><a:p>(<a:pPr[^>]*>(?:(?!</a:p>).)*?)?<a:r>(<a:rPr[^/]*/>[^<]*)<a:t>(to leverage)',
        r'\1 \4</a:t></a:r><a:endParaRPr/></a:p><a:p>\2<a:r>\3<a:t>',
        content, flags=re.DOTALL
    )

    return content


def fix_slide6(content):
    """Update subtitle to emphasize current students as PRIMARY audience."""
    targets = [
        'Know who you&#x2019;re talking to before you create content.',
        "Know who you\u2019re talking to before you create content.",
        "Know who you're talking to before you create content.",
    ]
    new_text = 'Current students are the primary audience for most COBE program marketing. Other audiences (prospective students, employers, partners) are secondary and typically handled by specialized teams with dedicated systems.'
    for old in targets:
        content = content.replace(old, new_text)
    return content


def fix_slide7(content):
    """Update subtitle to clarify these are RHM-specific examples."""
    content = content.replace(
        'Every piece of content should connect to at least one of these themes.',
        'These messaging pillars were developed for RHM certificate campaigns. Other programs can adapt these as examples, but they are not universal requirements.'
    )
    return content


def fix_slide8(content):
    """Add context that funnel is for Hands-On teams."""
    content = content.replace(
        'Match your goal to the right funnel stage and channel.',
        'Match your goal to the right funnel stage and channel. This framework is primarily for Hands-On teams. Hands-Off users can skip this \u2014 just tell COBE Marketing your goal and they will pick the right channels for you.'
    )
    return content


def fix_slide12(content):
    """Remove specific rows and channel references from the channel table."""
    # Slide 12 uses <a:tr> table rows for the channel recommendation table.

    # Remove table row containing "Recruit to certificate"
    content = re.sub(
        r'<a:tr[^>]*>(?:(?!</a:tr>).)*?Recruit to certificate(?:(?!</a:tr>).)*?</a:tr>',
        '',
        content, flags=re.DOTALL
    )

    # Remove table row containing "Build employer awareness"
    content = re.sub(
        r'<a:tr[^>]*>(?:(?!</a:tr>).)*?Build employer awareness(?:(?!</a:tr>).)*?</a:tr>',
        '',
        content, flags=re.DOTALL
    )

    # Remove "Class Visits" from text
    content = re.sub(r',?\s*Class Visits', '', content)
    content = re.sub(r'Class Visits,?\s*', '', content)
    content = re.sub(r'\s*\+\s*Class Visits', '', content)
    content = re.sub(r'Class Visits\s*\+\s*', '', content)

    # Remove "Paid Ads" from text
    content = re.sub(r',?\s*Paid Ads', '', content)
    content = re.sub(r'Paid Ads,?\s*', '', content)
    content = re.sub(r'\s*\+\s*Paid Ads', '', content)
    content = re.sub(r'Paid Ads\s*\+\s*', '', content)

    # Remove the podcast reference from subtitle (already renamed by global replace)
    content = content.replace(
        '(Includes Bronco Business Podcast for partnerships and thought leadership.)',
        ''
    )

    return content


def fix_slide13(content):
    """Remove 1 WEEK OUT, DAY OF, and 1 WEEK AFTER rows (shapes by y-position)."""
    # These are individual shapes, not table rows.
    # Shape positions:
    #   1 WEEK OUT row:  y=3246120
    #   DAY OF row:      y=3794760
    #   1 WEEK AFTER row: y=4343400
    # Remove all <p:sp> elements with y offsets in these positions
    # Also need to catch decorative background shapes for those rows

    # Approach: remove any <p:sp> whose <a:off> y value is >= 3246120 and < 4700000
    # (keeping footer at y=4754880 and page number)

    def remove_shapes_by_y(xml_content, y_min, y_max):
        """Remove <p:sp> elements where a:off y is between y_min and y_max."""
        result = xml_content
        # Find all <p:sp>...</p:sp> blocks
        shapes = list(re.finditer(r'<p:sp>(?:(?!</p:sp>).)*?</p:sp>', result, re.DOTALL))
        # Process in reverse order to maintain positions
        for m in reversed(shapes):
            shape_xml = m.group()
            off_match = re.search(r'<a:off x="\d+" y="(\d+)"', shape_xml)
            if off_match:
                y = int(off_match.group(1))
                if y_min <= y < y_max:
                    result = result[:m.start()] + result[m.end():]
        return result

    content = remove_shapes_by_y(content, 3246120, 4700000)
    return content


def fix_slide14(content):
    """Simplify sample request email - remove Channels requested, Target dates, Success metric."""
    # Remove lines/runs containing these field labels
    for pattern in [
        r'Channels [Rr]equested:?[^<]*',
        r'Target [Dd]ates?:?[^<]*',
        r'Success [Mm]etric:?[^<]*',
    ]:
        content = re.sub(pattern, '', content)
    return content


def create_consolidated_slide15(original_content):
    """Replace slide 15 content with consolidated channel overview."""
    # Change the title
    original_content = original_content.replace(
        'Hands-Off: Instagram',
        'Hands-Off Channel Overview'
    )

    # Replace the "BEST FOR" label and subtitle description
    original_content = original_content.replace(
        'Student engagement, event awareness, program visibility. 1,436 followers.',
        'A quick reference for all Hands-Off channels. Tell COBE your goal and provide the basics \u2014 they handle the rest.'
    )

    # Replace section labels and content with consolidated channel summaries
    # Using the EXACT text runs found in the original file
    original_content = original_content.replace(
        'ASSET TYPES',
        'INSTAGRAM'
    )
    original_content = original_content.replace(
        'Square posts (1080x1080), Stories/Reels (1080x1920)',
        'Best for student engagement, event awareness, program visibility. Provide: event details, photos (1080x1080 or 1080x1920), and key message.'
    )
    original_content = original_content.replace(
        'CADENCE',
        'LINKEDIN'
    )
    original_content = original_content.replace(
        '3-4 posts/week during peak recruitment; stories around events',
        'Best for employer partnerships, alumni highlights, professional storytelling. Provide: professional copy, speaker/employer info, and banner image.'
    )
    original_content = original_content.replace(
        'LEAD TIME',
        'EMAIL NEWSLETTER'
    )
    original_content = original_content.replace(
        '10 business days; urgent posts need marketing manager approval',
        'Best for driving event attendance, applications, program inquiries. Provide: rectangle landscape photo, short paragraph with value props, and CTA link.'
    )
    original_content = original_content.replace(
        'WHAT YOU SEND COBE',
        'INTERNAL CHANNELS'
    )
    original_content = original_content.replace(
        'Event/program details, photos (min 1080px), registration link, goal/success metric.',
        'Best for reaching current students via classroom slides and MBEB screens. Provide: event details and a 1920x1080 graphic.'
    )
    # Clear out the remaining detail sections
    original_content = original_content.replace(
        'PRO TIPS',
        ''
    )
    original_content = original_content.replace(
        'Reels outperform static posts for reach. Use Link Stickers in Stories. Always include a clear CTA.',
        ''
    )
    original_content = original_content.replace(
        'FUNNEL STAGE',
        ''
    )
    original_content = original_content.replace(
        'Awareness + Consideration. Best for current students and prospective undergrads.',
        ''
    )

    return original_content


def fix_slide19(content):
    """Add note that podcast is a special case requiring Kate Watson coordination."""
    # The text "Employer partnerships, alumni engagement, thought leadership, student recruiting influence."
    # is in its own shape/text run
    content = content.replace(
        'Employer partnerships, alumni engagement, thought leadership, student recruiting influence.',
        'Employer partnerships, alumni engagement, thought leadership, student recruiting influence. NOTE: The Bronco Business Podcast is a special case that COBE does not generally offer broadly. Direct coordination with Kate Watson is required.'
    )
    return content


def fix_slide20(content):
    """Add special case note to podcast process slide."""
    content = content.replace(
        'Identify the External Guest',
        'Identify the External Guest (coordinate directly with Kate Watson)'
    )
    return content


def fix_slide25(content):
    """Add note about Kim's calendar coordination."""
    content = content.replace(
        'Submit completed calendar with assets to cobe-info',
        'Submit completed calendar with assets to cobe-info. Note: Submitted calendars will be coordinated with Kim and the main COBE social calendar. Your schedule must dovetail with the existing COBE posting calendar.'
    )
    return content


def fix_slide26(content):
    """Remove class visits from budget slide."""
    content = content.replace(', class visits,', ',')
    content = content.replace(', class visits)', ')')
    content = content.replace('class visits, ', '')
    content = content.replace(', class visits', '')
    return content


def fix_slide29(content):
    """Change Square-ish photo to Rectangle landscape photo."""
    content = content.replace('Square-ish photo', 'Rectangle landscape photo')
    content = content.replace('square-ish photo', 'rectangle landscape photo')
    return content


def fix_slide31(content):
    """Update WordPress section for specialized teams only."""
    content = content.replace(
        'WORDPRESS ACCESS REQUIREMENTS',
        'WORDPRESS ACCESS (Specialized Teams Only)'
    )
    content = content.replace(
        'WEB UPDATE REQUEST FORMAT',
        'NOTE: Most programs should submit web update requests to cobe-info@boisestate.edu rather than pursuing WordPress access directly. The university is reducing the number of WordPress admins and webpages. | WEB UPDATE REQUEST FORMAT'
    )
    return content


def fix_slide39(content):
    """Fix @boisestatecobe reference in social posting context."""
    content = content.replace(
        'Social posting to @boisestatecobe accounts',
        'Social posting to COBE accounts (@boisestatecobe on Instagram; College of Business and Economics at Boise State on LinkedIn)'
    )
    return content


def fix_slide47(content):
    """Fix @boisestatecobe reference."""
    content = content.replace(
        'Social posts published to @boisestatecobe accounts',
        'Social posts published to COBE accounts (@boisestatecobe on Instagram; College of Business and Economics at Boise State on LinkedIn)'
    )
    return content


def fix_slide50(content):
    """Simplify tracking guidance."""
    content = content.replace(
        'Track with Flowcode + UTMs + Looker Studio',
        'Track with QR Codes (everyone) + UTMs + Looker Studio (specialized teams)'
    )
    content = content.replace(
        'Flowcode QR per physical placement. UTM parameters on every digital link. Looker Studio dashboard compares channels side-by-side.',
        'For general users: Create Flowcode QR codes for your physical placements \u2014 this is the main tracking most programs need. For specialized teams: Add UTM parameters on every digital link and use Looker Studio dashboards to compare channels side-by-side.'
    )
    return content


def fix_slide56_contacts(content):
    """Update contacts table: remove McKenna, Ann, Ashley, Kent; add Central Marketing."""
    # Slide 56 uses <a:tr> table rows for the contacts table.

    # Remove table row containing "McKenna Howard"
    content = re.sub(
        r'<a:tr[^>]*>(?:(?!</a:tr>).)*?McKenna Howard(?:(?!</a:tr>).)*?</a:tr>',
        '',
        content, flags=re.DOTALL
    )

    # Remove table row containing "Ann Hottinger"
    content = re.sub(
        r'<a:tr[^>]*>(?:(?!</a:tr>).)*?Ann Hottinger(?:(?!</a:tr>).)*?</a:tr>',
        '',
        content, flags=re.DOTALL
    )

    # Remove table row containing "Ashley Walter"
    content = re.sub(
        r'<a:tr[^>]*>(?:(?!</a:tr>).)*?Ashley Walter(?:(?!</a:tr>).)*?</a:tr>',
        '',
        content, flags=re.DOTALL
    )

    # Remove table row containing "Kent Neupert"
    content = re.sub(
        r'<a:tr[^>]*>(?:(?!</a:tr>).)*?Kent Neupert(?:(?!</a:tr>).)*?</a:tr>',
        '',
        content, flags=re.DOTALL
    )

    # Add Central Marketing reference - modify Laura Chiuppi's responsibility text
    # The original text (after global replace) will have "External engagement, Bronco Business Podcast coordination"
    # or similar. Replace with cleaned up version + Central Marketing info.
    content = content.replace(
        'External engagement',
        'External engagement | For design support: Boise State Central Marketing (boisestate.edu/communicationsandmarketing/creative/request-design-support/)',
        1  # Only first occurrence
    )

    return content


def fix_slide57_brand(content):
    """Update brand reference to point to brand website."""
    content = content.replace(
        'KEY GUIDELINES',
        'KEY GUIDELINES (For complete visual identity including print color profiles, visit boisestate.edu/brand/visual-identity/)'
    )
    content = content.replace(
        'BOISE STATE COLORS',
        'BOISE STATE COLORS (Digital \u2014 see brand site for print versions)'
    )
    return content


def fix_slide59_faq(content):
    """Fix FAQ formatting: indent answers, remove Marq, fix Resource Library reference."""
    # Add arrow prefix to answers
    content = content.replace(
        '>No. All content goes through cobe-info for review.<',
        '>\u2192 No. All content goes through cobe-info for review.<'
    )
    content = content.replace(
        '>Standard is 10 business days. Rush = director approval.<',
        '>\u2192 Standard is 10 business days. Rush = director approval.<'
    )
    content = content.replace(
        '>Hands-Off: COBE creates them. Hands-On: use templates.<',
        '>\u2192 Hands-Off: COBE creates them. Hands-On: use Canva templates (replacing Marq, campus-wide transition).<'
    )
    content = content.replace(
        '>Yes! Email cobe-info, subject: STUDENT ORG REQUEST.<',
        '>\u2192 Yes! Email cobe-info, subject: STUDENT ORG REQUEST.<'
    )
    content = content.replace(
        '>See the Resource Library slide or the playbook repository.<',
        '>\u2192 See the Appendix or contact cobe-info@boisestate.edu for templates.<'
    )
    content = content.replace(
        'Resource Library slide',
        'Appendix'
    )
    content = content.replace(
        '>Contact Kate Watson directly for urgent changes or crisis situations.<',
        '>\u2192 Contact Kate Watson directly for urgent changes or crisis situations.<'
    )

    # Remove Marq references
    content = content.replace('Marq templates', 'Canva templates (campus-wide transition in progress)')
    content = content.replace('Marq', 'Canva')

    return content


def fix_slide61_changelog(content):
    """Add new changelog entry for v4.0."""
    content = content.replace(
        'Final audit: corrected page numbers, fixed data tables, updated changelog, distribution-ready',
        'Final audit: corrected page numbers, fixed data tables, updated changelog, distribution-ready | v4.0 \u2014 2026-02-23 \u2014 Ben Howerton \u2014 Kate Watson revisions: consolidated hands-off channels, fixed audience focus, renamed Bronco Business Podcast, simplified for department chairs, updated contacts and brand references.'
    )
    return content


def fix_presentation_xml(content):
    """Remove deleted slide references from presentation.xml."""
    for rid in DELETED_RIDS:
        content = re.sub(rf'<p:sldId\s+id="\d+"\s+r:id="{rid}"\s*/?>', '', content)
    return content


def fix_presentation_rels(content):
    """Remove deleted slide relationships from presentation.xml.rels."""
    for rid in DELETED_RIDS:
        content = re.sub(rf'<Relationship\s+Id="{rid}"[^/]*/?>', '', content)
    return content


def fix_content_types(content):
    """Remove deleted slide/notes entries from [Content_Types].xml."""
    for s in SLIDES_TO_DELETE:
        content = re.sub(rf'<Override[^>]*slide{s}\.xml[^/]*/?>', '', content)
        content = re.sub(rf'<Override[^>]*notesSlide{s}\.xml[^/]*/?>', '', content)
    return content


def update_page_numbers(content, old_page, new_page):
    """Update the displayed page number on a slide."""
    if old_page == new_page:
        return content

    old_str = str(old_page)
    new_str = str(new_page)

    # Page number is in a shape at x=8412480, y=4754880 containing just the number
    pattern = (
        r'(<a:off x="8412480" y="4754880"/>.*?)'
        r'<a:t>' + re.escape(old_str) + r'</a:t>'
    )
    content = re.sub(pattern, r'\1<a:t>' + new_str + '</a:t>', content, count=1, flags=re.DOTALL)
    return content


def compute_new_page_numbers():
    """Map slide file numbers to (old_displayed_page, new_displayed_page)."""
    return {
        19: (19, 16), 20: (20, 17), 21: (21, 18),
        23: (23, 19), 24: (24, 20), 25: (25, 21),
        26: (26, 22), 27: (27, 23), 28: (28, 24),
        29: (29, 25), 30: (30, 26), 31: (31, 27),
        32: (32, 28),
        34: (33, 29), 35: (34, 30), 36: (35, 31),
        38: (36, 32), 39: (37, 33), 40: (38, 34),
        41: (39, 35), 42: (40, 36), 43: (41, 37),
        44: (42, 38), 45: (43, 39), 46: (44, 40),
        47: (45, 41), 48: (46, 42), 49: (47, 43),
        50: (48, 44), 51: (49, 45),
        53: (51, 46), 54: (52, 47),
        56: (54, 48), 57: (55, 49), 58: (56, 50),
        59: (57, 51), 60: (58, 52), 61: (59, 53),
    }


def main():
    renumber_map = compute_new_page_numbers()

    with zipfile.ZipFile(INPUT_PATH, 'r') as zin:
        with zipfile.ZipFile(OUTPUT_PATH, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in FILES_TO_SKIP:
                    print(f"  SKIPPING (deleted): {item.filename}")
                    continue

                data = zin.read(item.filename)

                if item.filename.endswith('.xml') or item.filename.endswith('.rels'):
                    content = data.decode('utf-8')

                    # Global replacements
                    content = global_text_replacements(content, item.filename)

                    # Slide-specific fixes
                    slide_fixes = {
                        'ppt/slides/slide1.xml': fix_slide1,
                        'ppt/slides/slide6.xml': fix_slide6,
                        'ppt/slides/slide7.xml': fix_slide7,
                        'ppt/slides/slide8.xml': fix_slide8,
                        'ppt/slides/slide12.xml': fix_slide12,
                        'ppt/slides/slide13.xml': fix_slide13,
                        'ppt/slides/slide14.xml': fix_slide14,
                        'ppt/slides/slide15.xml': create_consolidated_slide15,
                        'ppt/slides/slide19.xml': fix_slide19,
                        'ppt/slides/slide20.xml': fix_slide20,
                        'ppt/slides/slide25.xml': fix_slide25,
                        'ppt/slides/slide26.xml': fix_slide26,
                        'ppt/slides/slide29.xml': fix_slide29,
                        'ppt/slides/slide31.xml': fix_slide31,
                        'ppt/slides/slide39.xml': fix_slide39,
                        'ppt/slides/slide47.xml': fix_slide47,
                        'ppt/slides/slide50.xml': fix_slide50,
                        'ppt/slides/slide56.xml': fix_slide56_contacts,
                        'ppt/slides/slide57.xml': fix_slide57_brand,
                        'ppt/slides/slide59.xml': fix_slide59_faq,
                        'ppt/slides/slide61.xml': fix_slide61_changelog,
                    }

                    if item.filename in slide_fixes:
                        content = slide_fixes[item.filename](content)

                    # Structural fixes
                    if item.filename == 'ppt/presentation.xml':
                        content = fix_presentation_xml(content)
                    elif item.filename == 'ppt/_rels/presentation.xml.rels':
                        content = fix_presentation_rels(content)
                    elif item.filename == '[Content_Types].xml':
                        content = fix_content_types(content)

                    # Page number updates
                    if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
                        match = re.search(r'slide(\d+)\.xml', item.filename)
                        if match:
                            slide_num = int(match.group(1))
                            if slide_num in renumber_map:
                                old_page, new_page = renumber_map[slide_num]
                                content = update_page_numbers(content, old_page, new_page)

                    data = content.encode('utf-8')

                zout.writestr(item, data)

    print(f"\nOutput saved to: {OUTPUT_PATH}")
    print(f"Original size: {os.path.getsize(INPUT_PATH):,} bytes")
    print(f"New size: {os.path.getsize(OUTPUT_PATH):,} bytes")


if __name__ == '__main__':
    main()
