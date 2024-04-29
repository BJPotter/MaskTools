import pandas as pd
import re 
import os

#指定列进行加密处理
def mask_phone_v1(value):
    return value[:3]+'****'
    
#按照正则匹配进行加密处理
def mask_phone_v2(value):
    """
    将手机号替换为****
    """
    if pd.isnull(value):  # 新增判断，如果值为空则直接返回
        return value
    value = str(value)
    #手机号、地址的正则表达式
    phone_pattern = r'(13\d{9}|14[5|7]\d{8}|15\d{9}|166{\d{8}|17[3|6|7]{\d{8}|18\d{9})'
    address_pattern = r'(中国北京|号楼|单元|街道)'
    
    # 通过正则匹配手机号和地址
    # 非手机号地址，返回原值
    if re.findall(phone_pattern, value) or re.findall(address_pattern, value) :  
        return value[:3]+'****'
    return value  

def excel_encrypt_file(file_path, save_path):
    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        #标题所在行，注意所在行
        df = pd.read_excel(file_path, header=1, keep_default_na=False)
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path, header=0, keep_default_na=False)
    else:
        print("无法识别的文件格式。")
        return
#<<<<<<<<<<<<<<<<<<<<<<<
    #指定处理，开放下面这两行的注释
    #不处理的列
    #skip_columns = ["账号", "宽带账号","account"]
    #必须加密的列
    #necessary_columns = ["CRM联系电话", "融合手机号","地址","标准小区"]
#----------------------------
    #默认处理开放下面这两行处理
    skip_columns = []
    necessary_columns = []
#>>>>>>>>>>>>>>>>>>>>>
    if (len(skip_columns)== 0 and len(necessary_columns) == 0):
        for col in df.columns:
            df[col] = df[col].astype(str)
            df[col] = df[col].apply(mask_phone_v2)
    else:   
        for col in df.columns:
            if col not in skip_columns:
                if col in necessary_columns:
                    df[col] = df[col].astype(str)
                    df[col] = df[col].apply(mask_phone_v1)
                df[col] = df[col].astype(str)
                df[col] = df[col].apply(mask_phone_v2)

    _, ext = os.path.splitext(file_path)

    with pd.ExcelWriter(save_path, engine='xlsxwriter', options={'strings_to_urls': False}) as writer:
        df.to_excel(writer, index=False)  # 保留原来的 index




# 调用函数处理文件。
#excel_encrypt_file('D:\Desktop\VIP0627.xlsx', 'D:\Desktop\VIP0627-mask.xlsx')

