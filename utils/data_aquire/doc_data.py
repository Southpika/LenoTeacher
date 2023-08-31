from docx import Document

def handle_bold(paragraph):
    bold_text = ''
    normal_text = ''
    for run in paragraph.runs:  
        if run.bold:      
            bold_text += run.text
        else:
            normal_text += run.text         
    return bold_text.replace('\u3000',' ').strip(),normal_text.replace('\u3000',' ').strip()

def convert_paragraph_to_markdown(paragraph):
    b_t,n_t = handle_bold(paragraph)
    markdown_text = ""
    if b_t: 
        markdown_text +=  f"\n# {b_t} \n"
    if n_t: markdown_text += n_t
    return markdown_text

def get_docx_markdown(doc_path = r"D:\tzh\人教版思想政治\01 必修一\PPT课件+教案\第二课 只有社会主义才能救中国\2.2 社会主义制度在中国的确立教学设计.docx"):
    document = Document(doc_path)

    markdown_output = ""

    for paragraph in document.paragraphs:
        output = convert_paragraph_to_markdown(paragraph)
        if output: 
            if output[0] != '#' and markdown_output[-1:] != '\n':
                output = '\n' + output
            markdown_output += output

    return markdown_output.strip('\n')

if __name__ == '__main__':
    doc_path = input('请输入地址：')
    if not doc_path:
        print(get_docx_markdown())
    else:
        print(get_docx_markdown(doc_path))