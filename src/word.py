import re
from docx import Document
def mask_content(text):
    """
    检查文本是否包含电话号码或地址，如果存在，则将其替换为"****"
    """
    # 定义匹配电话号码和地址的正则表达式
    phone_pattern = r'(13\d{9}|14[5|7]\d{8}|15\d{9}|166{\d{8}|17[3|6|7]{\d{8}|18\d{9})'
    address_pattern = r'(中国北京|号楼|单元)'

    if re.findall(phone_pattern,text):
        #print(text)
        return re.sub(phone_pattern, text[:3]+'****', text)
    elif re.findall(address_pattern,text):
        #print(text)
        return text[:5]+'****'+text[-1]
    return text


def word_encrypt_file(file_path, save_path):
    """
    读取Word文档，找到可能的电话号码和地址，并替换为"****"，然后保存到新的文件中
    """
    # 打开文档
    doc = Document(file_path)
    # 对每个段落进行处理
    for para in doc.paragraphs:
        sentences = re.split('([。？！;；,，：:])', para.text)
        sentences = [sentences[i]+sentences[i+1] for i in range(0, len(sentences)-1, 2)]
        for sentence in sentences:
            masked_sentence = mask_content(sentence)  # 对每个句子做处理
            para.text = para.text.replace(sentence, masked_sentence)  # 把原句子替换成处理后的句子
    # 保存处理后的文档
    doc.save(save_path)

