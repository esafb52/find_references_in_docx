from docx import Document


def save_references_file(txt_file_name, lst_txt):
    my_doc = open(txt_file_name, 'w', encoding='utf-8')
    for p in lst_txt:
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
                if 8 < len(temp_txt) < 100 and has_numbers(temp_txt):
                    ls_txt_data.append(temp_txt.strip()[1:-1])
                temp_txt = ''
    ls = list(set(ls_txt_data))
    ls.sort()
    for item in ls:
        print(item)
    return ls


def has_numbers(input_str):
    return any(char.isdigit() for char in input_str)
    

if __name__ == '__main__':
    docx_file_path = 'g:/qq.docx'
    result_file_path = 'g:/final_reference.txt'
    lst_ref = find_references(docx_file_path)
    save_references_file(result_file_path, lst_ref)
    print('find references complete  !!!!!!')
