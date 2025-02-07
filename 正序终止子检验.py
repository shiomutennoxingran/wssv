import xlrd
import string
import re

sq = 120  # 选取的开头结尾多少碱基为保守区域
data = xlrd.open_workbook(r'C:/Users/lizhouquan/Desktop/论文返工/终止子检验.xls')  # 读取表格文件
table = data.sheets()[0]
cols_list0 = table.col_values(colx=0)  # 读取需要检验的orf的名称
#cols_list1 = table.col_values(colx=1)  # 读取需要寻找的orf的名称

for j in range(0, len(cols_list0)):
    j_str = ''  # 清除j_str内字符串
    j_str = str(cols_list0[j])  # 读取当前orf名称
    filename = 'C:/Users/lizhouquan/Desktop/论文返工/再更新orf/' + j_str + '.txt'  # 组成读取待检验orf文件的路径
    filename1 = 'C:/Users/lizhouquan/Desktop/论文返工/再更新检验/正序/' + j_str + '.txt'  # 组成保存检验完成的orf的路径
    with open(filename, encoding='gbk') as f1: #utf-8 gbk
        str0 = f1.read()
        str1 = str0[str0.index('\n') + 1:len(str0)]
        str1 = str1.replace('\n', '')  # 得到fasta格式>的内容
        #print(str1)
        #str2 = str1[0:sq]  # 需要查寻的ORF的开头保守区域
        #str3 = str1[len(str1) - sq:len(str1)]  # 需要查寻的ORF的结尾保守区域
    with open(filename1, "w") as d:  # 创建保存结果的txt
        a = str(len(str1))  # orf的长度
        print(j_str)#orfx
        #d.write('该orf的标准长度为' + a + '\n')  # 在最开头输入该orf的长度方便校对

        # 使用正则表达式找到每个 '>' 的位置
        str1 = str1 +'>' #方便识别而增加
        match_positions = [match.start() for match in re.finditer(r'>', str1)]
        # 初始化一个空列表来存储结果
        results = []
        # 遍历匹配位置，提取两个 '>' 之间的内容
        for i in range(len(match_positions) - 1):
            start_pos = match_positions[i] + 1
            end_pos = match_positions[i + 1]
            sequence_part = str1[start_pos:end_pos] if end_pos != -1 else str1[start_pos:]
            filtered_sequence = ''.join([char for char in sequence_part if char in 'ATCG找不到捏'])
            if filtered_sequence:
                results.append(filtered_sequence)
        # 打印结果
        #print(results)
        n = 1
        for i in range(0, len(results), 1):
            group_chunks = results[i:i+1]  
            group = ''.join(group_chunks) if group_chunks else ''
            #继续拆分group
            split_groups = []
            # 循环遍历字符串，每次取三个字符
            for i in range(0, len(group), 3):
            # 使用切片获取当前三个字符的组
                three_char_group = group[i:i+3]
            # 将组添加到列表中
                split_groups.append(three_char_group)
                #print(split_groups)
            split_groups.pop()
            print(split_groups)
            #查找终止子
            target_substrings = ["TAG", "TAA", "TGA", "找不到"]
            contains_target = "不含有终止子，是orf"
            for group in split_groups:
                if any(substring in group for substring in target_substrings):
                    contains_target = "含有终止子，不认为是orf/缺失"
            put = str(n)+ contains_target + '\n'
            d.write(put)
            print(put)
            n = n+1