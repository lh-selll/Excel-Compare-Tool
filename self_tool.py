import sys
import re
import copy

def process_text(text, Min_length):
    text_temp = copy.copy(text).strip()
    if len(text_temp.split(",")) >= Min_length:
        text_temp = text_temp.replace(" ", "")
        text_temp = text_temp.split(',')
        # del text_temp[0]
    elif len(text_temp.split("0X")) >= Min_length+1:
        text_temp = text_temp.replace(" ", "")
        text_temp = text_temp.split('0X')
        del text_temp[0]
    elif len(text.split("0x")) >= Min_length+1:
        text_temp = text_temp.replace(" ", "")
        text_temp = text_temp.split('0x')
        del text_temp[0]
    elif len(re.split('\s+', text_temp, 0)) >= Min_length:
        text_temp =  re.split('\s+', text_temp, 0)
    else:
        try:
            print(f"text_temp = {text_temp}")
            text_temp = text_temp.replace('_x000D_', '').replace('\r', '').replace('\n', '').replace(' ', '').replace("0x", "").replace("0X", "")
            print(f"text_temp = {text_temp}")
            text_obj = ""
            # 将十六进制字符串转换为字节对象
            byte_obj = bytes.fromhex(text_temp)
            # 将字节对象转换为包含单个字节的列表
            # 转换为十六进制数列表：借助列表推导式 [hex(byte)[2:].zfill(2) for byte in byte_obj] 来遍历字节对象中的每个字节。
            # hex(byte)：把字节转换为十六进制字符串，结果会带有 0x 前缀。
            # [2:]：去掉 0x 前缀。
            # .zfill(2)：确保每个十六进制字符串都是两位，若不足两位则在前面补 0。
            byte_result = [hex(byte)[2:].zfill(2) for byte in byte_obj]
            return byte_result
        except ValueError as e:
            print(f"process_text() , 无效的十六进制字符: {e}\n可能包含非十六进制字符，或者字符数量为奇数")
            return text_temp
    
    print(f"process_text = {text_temp}")
    return text_temp

if __name__ == "__main__":
    # 主程序入口，创建应用实例并显示图形界面。
    while 1:
        try:
            string = input("input your string： ")
            length = int(input("input min length of your string: "))
            print(f"process_text = {process_text(string, length)}")
        except ValueError:
            error = f"Seed value inputted is invalid"
            print(error)
        