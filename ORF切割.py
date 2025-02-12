import xlrd

sq = 12  # 选取的开头结尾保守区域长度
data = xlrd.open_workbook(r'C:/Users/lizhouquan/Desktop/论文返工/序列.xls')
table = data.sheets()[0]
cols_list0 = table.col_values(colx=0)  # 未知序列名称
cols_list1 = table.col_values(colx=1)  # ORF名称

for j in range(len(cols_list1)):
    j_str = str(cols_list1[j])
    filename = f'C:/Users/lizhouquan/Desktop/论文返工/标准orf/{j_str}.txt'
    filename1 = f'C:/Users/lizhouquan/Desktop/论文返工/待查看/{j_str}.txt'
    
    with open(filename, encoding='utf-8') as f1:
        str0 = f1.read()
        str1 = str0[str0.index('\n')+1:].replace('\n', '')
        str2 = str1[:sq]  # 标准开头
        str3 = str1[-sq:]  # 标准结尾

    with open(filename1, "w") as d:
        d.write(f'该orf的标准长度为{len(str1)}\n')

        for i in range(5):  # 根据实际情况调整循环次数
            i_str = str(cols_list0[i])
            filename2 = f'C:/Users/lizhouquan/Desktop/对齐后新基因组/{i_str}.txt'
            
            with open(filename2, encoding='utf-8') as f2:
                strd = f2.read().replace('\n', '')
                try:
                    # 查找保守区域
                    start_idx = strd.index(str2)
                    strd1 = strd[start_idx:]
                    end_idx = strd1.index(str3) + len(str3)
                    found_orf = strd[start_idx:start_idx+end_idx]
                    standard_orf = str1

                    # 计算整体相似度
                    matches = sum(f == s for f, s in zip(found_orf, standard_orf))
                    similarity = matches/len(standard_orf)*100

                    # 计算两端相似度
                    found_start = found_orf[:sq]
                    found_end = found_orf[-sq:]
                    
                    start_matches = sum(f == s for f, s in zip(found_start, str2))
                    start_similarity = start_matches/sq*100
                    
                    end_matches = sum(f == s for f, s in zip(found_end, str3))
                    end_similarity = end_matches/sq*100

                    # 写入结果
                    d.write(
                        f">{i_str} "
                        f"(长度:{len(found_orf)}) "
                        f"(起始:{start_idx}) "
                        f"(结束:{start_idx+len(found_orf)}) "
                        f"整体:{similarity:.2f}% "
                        f"开头:{start_similarity:.2f}% "
                        f"结尾:{end_similarity:.2f}%\n"
                    )
                    d.write(found_orf + "\n\n")

                except ValueError:
                    d.write(f">{i_str}\n找不到保守区域\n\n")