from docx import Document


def save_references_file(txt_file_name, lst_txt):
    if lst_txt is not None:
        if len(lst_txt) > 1:
            lst_txt.sort()
    ls_final = set(lst_txt)
    my_doc = open(txt_file_name, 'w', encoding='utf-8')
    for p in ls_final:
        my_doc.writelines(str(p).strip() + '\n')
    my_doc.close()


def find_references(path):
    ls_txt_data = []
    mode = False
    temp_txt = ''
    document = Document(path)
    for para in document.paragraphs:
        all_para_txt = str(para.text)
        for char in all_para_txt:
            if char == '(':
                mode = True
            if mode:
                temp_txt = temp_txt + char
            if char == ')':
                mode = False
                if 7 < len(temp_txt) < 100:
                    ls_txt_data.append(temp_txt.strip())
                    print(temp_txt[1:-1])
                temp_txt = ''
    return ls_txt_data


def main():
    docx_file_path = 'g:/qq.docx'
    result_file_path = 'g:/final_reference.txt'
    lst_ref = find_references(docx_file_path)
    save_references_file(result_file_path, lst_ref)
    print('find references complete  !!!!!!')
    

if __name__ == '__main__':
    main()
   
