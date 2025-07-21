from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
import tempfile, os

app = Flask(__name__)
CORS(app)

def run_to_html(run):
    raw_text = run.text.strip()

    # ✅ Detect and pass through raw HTML tags (img/iframe)
    if raw_text.startswith('<img') or raw_text.startswith('<iframe'):
        return raw_text

    if run._element.xpath('.//w:hyperlink'):
        link = run._element.xpath('.//w:hyperlink')[0]
        r_id = link.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        hyperlink = run.part.related_parts[r_id].target_ref if r_id else None
        text = run.text.strip()
        if hyperlink:
            return f'<a href="{hyperlink}" target="_blank">{text}</a>'

    text = run.text
    if not text.strip():
        return ''

    # Escape HTML characters, but only after newline check
    text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    # Only replace \n with <br> if not at the start
    if '\n' in text:
        lines = text.split('\n')
        text = '<br>'.join(line.strip() for line in lines if line.strip())

    styles = []
    if run.bold:
        styles.append("font-weight:bold;")
    else:
        styles.append("font-weight:normal;")

    if run.italic:
        styles.append("font-style:italic;")

    if run.font.size:
        try:
            styles.append(f"font-size:{run.font.size.pt:.2f}pt;")
        except:
            pass

    if run.font.name:
        styles.append(f"font-family:'{run.font.name}';")

    if run.font.color and run.font.color.rgb:
        styles.append(f"color:#{run.font.color.rgb};")

    style_str = ''.join(styles)
    return f'<span style="{style_str}">{text}</span>'


def detect_list_type(para):
    """
    Returns (is_list, list_type) where list_type is 'ul' or 'ol'
    """
    style_name = para.style.name.lower()
    first_text = para.text.strip()

    # ✅ Strong detection of unordered bullet characters
    bullet_symbols = ['•', '-', '‣', '◦', '▪']
    if first_text and first_text[0] in bullet_symbols:
        return True, 'ul'

    if 'bullet' in style_name:
        return True, 'ul'
    if 'number' in style_name:
        return True, 'ol'

    # Fallback to Word structure
    numPr = para._p.pPr.numPr if para._p.pPr is not None and para._p.pPr.numPr is not None else None
    if numPr:
        ilvl = numPr.ilvl.val if numPr.ilvl is not None else 0
        if ilvl == 0:
            return True, 'ol'
        else:
            return True, 'ul'

    return False, None

def para_to_html(para):
    html = ''.join([run_to_html(run) for run in para.runs])
    if not html.strip():
        return '', None

    # Detect list
    is_list, list_type = detect_list_type(para)

    color_style = ''
    if para.runs:
        first = para.runs[0]
        if first.font.color and first.font.color.rgb:
            color_style = f'style="color:#{first.font.color.rgb};"'

    if is_list and list_type:
        return f"<li {color_style}>{html}</li>", list_type

    return f"<p>{html}</p>", None

def wrap_list(items, list_type):
    if not list_type or not items:
        return ''.join(items)
    return f"<{list_type}>{''.join(items)}</{list_type}>"

def docx_to_html_sections(doc_path):
    doc = Document(doc_path)
    current_section = 0  # 0=head, 1=text, 2=faq
    sections = [[], [], []]  # head, text, faq_raw

    list_buffer = []
    list_type = None

    for para in doc.paragraphs:
        raw_text = para.text.strip()
        if not raw_text and not para.runs:
            continue

        if raw_text == '#####':
            if list_buffer and current_section == 1:
                sections[1].append(wrap_list(list_buffer, list_type))
                list_buffer = []
                list_type = None
            current_section += 1
            continue

        html, detected_list_type = para_to_html(para)
        if not html:
            continue

        if current_section <= 1:
            if detected_list_type:
                if list_type != detected_list_type and list_buffer:
                    sections[1].append(wrap_list(list_buffer, list_type))
                    list_buffer = []
                list_type = detected_list_type
                list_buffer.append(html)
            else:
                if list_buffer:
                    sections[1].append(wrap_list(list_buffer, list_type))
                    list_buffer = []
                    list_type = None
                sections[current_section].append(html)
        elif current_section == 2:
            is_bold = any(run.bold for run in para.runs if run.text.strip())
            sections[2].append({
                'type': 'question' if is_bold else 'answer',
                'html': html
            })

    if list_buffer and current_section == 1:
        sections[1].append(wrap_list(list_buffer, list_type))

    # Group FAQs
    faq_pairs = []
    current_q = None
    for item in sections[2]:
        if item['type'] == 'question':
            if current_q:
                faq_pairs.append({'question': current_q, 'answer': ''})
            current_q = item['html']
        elif item['type'] == 'answer':
            if current_q:
                faq_pairs.append({'question': current_q, 'answer': item['html']})
                current_q = None
    if current_q:
        faq_pairs.append({'question': current_q, 'answer': ''})

    return {
        'head': ''.join(sections[0]).strip(),
        'text': ''.join(sections[1]).strip(),
        'faq': faq_pairs
    }

@app.route('/parse-docx', methods=['POST'])
def handle_upload():
    if 'file' not in request.files:
        return jsonify({'error': 'File missing'}), 400

    uploaded = request.files['file']
    if not uploaded.filename.endswith('.docx'):
        return jsonify({'error': 'Only .docx allowed'}), 400

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    try:
        uploaded.save(tmp.name)
        tmp.close()
        result = docx_to_html_sections(tmp.name)
    except Exception as e:
        return jsonify({'error': f'Failed to parse: {str(e)}'}), 500
    finally:
        if os.path.exists(tmp.name):
            os.remove(tmp.name)

    return jsonify(result)

if __name__ == '__main__':
    app.run(port=5000, debug=True)