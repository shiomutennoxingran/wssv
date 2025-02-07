import xlrd
import string

sq = 12  # 选取的开头结尾多少碱基为保守区域
data = xlrd.open_workbook(r'C:/Users/lizhouquan/Desktop/论文返工/序列.xls')  # 读取表格文件
table = data.sheets()[0]
cols_list0 = table.col_values(colx=0)  # 读取所有未知序列名称
cols_list1 = table.col_values(colx=1)  # 读取需要寻找的orf的名称

for j in range(0, len(cols_list1)):
    j_str = str(cols_list1[j])  # 读取当前orf名称
    filename = 'C:/Users/lizhouquan/Desktop/论文返工/标准orf/' + j_str + '.txt'  # 标准orf路径
    filename1 = 'C:/Users/lizhouquan/Desktop/论文返工/待查看/' + j_str + '.txt'  # 结果保存路径
    
    with open(filename, encoding='utf-8') as f1:
        str0 = f1.read()
        str1 = str0[str0.index('\n') + 1:].replace('\n', '')  # 标准orf纯碱基序列
        str2 = str1[:sq]  # 开头保守区域
        str3 = str1[-sq:]  # 结尾保守区域
    
    with open(filename1, "w") as d:
        d.write(f'该orf的标准长度为{len(str1)}\n')  # 写入标准长度

        for i in range(0, 5):  # 遍历未知序列（根据实际情况调整循环次数）
            i_str = str(cols_list0[i])  # 当前序列名称
            filename2 = 'C:/Users/lizhouquan/Desktop/对齐后新基因组/' + i_str + '.txt'
            
            with open(filename2, encoding='utf-8') as f2:
                strd = f2.read().replace('\n', '')  # 读取未知序列
                try:
                    start_idx = strd.index(str2)
                    strd1 = strd[start_idx:]  # 截取开头之后的部分
                    end_idx = strd1.index(str3) + len(str3)  # 在截取部分找结尾
                    found_orf = strd[start_idx:start_idx+end_idx]  # 完整orf序列
                    standard_orf = str1  # 标准orf序列
                    
                    # 计算相似度
                    similarity = 0.0
                    #if len(found_orf) == len(standard_orf):
                    matches = sum(f == s for f, s in zip(found_orf, standard_orf))
                    similarity = matches / len(standard_orf) * 100
                    
                    # 写入结果
                    d.write(f'>{i_str}')
                    d.write(f'(长度:{end_idx})')
                    d.write(f'(起始:{start_idx})')
                    d.write(f'(结束:{start_idx+end_idx})')
                    d.write(f' 相似度:{similarity:.2f}%')
                    #d.write(' (达标)\n' if similarity >= 95 else ' (不达标)\n')
                    d.write(found_orf + '\n')
                
                except ValueError:
                    d.write(f'>{i_str}\n找不到保守区域\n')