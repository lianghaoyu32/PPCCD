import pandas as pd
import csv
import os
import re

'''读取文件名函数'''


def file_name(file_dir):
    for root, dirs, files in os.walk(file_dir):
        return files  # 当前文件夹下的所有文件的名称


'''获取文件夹名字的函数'''


def getfilename(filename):
    for root, dirs, files in os.walk(filename):
        array = dirs  # 当前文件夹下的所有子文件夹的名称
        if array:
            return array


# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 1/P2017004-P117-2-T5-1-20200217/'
# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 1/P2017004-P117-2-T5-2-20200217/'
# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 2/117-3-T5-1/'
# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 2/117-3-T5-2/'
file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 2/117-3-T5-3/'

# file_b = 'Z:/server/lhy/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 1/'
file_b = 'Z:/server/lhy/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 2/'


'''获取文件夹名字,生成列表并进行遍历'''
big_file = getfilename(file_a)

'''文件夹个数(文件夹列表中的序号)'''
big_file_num = 0

for big_file_name in big_file:
    print("循环文件夹")
    print(big_file_name)
    '''去掉DATA字符，只留下前几位字符'''
    '''文件路径'''

    csv_l = file_name(file_a + big_file_name + '/' + big_file_name + '/')
    csv_file_name = []  # 储存csv文件

    csv_pattern = re.compile('.+csv')  # 在文件夹中找出csv文件
    # panel = 1
    panel = 2

    for csv_i in csv_l:
        csv_result = csv_pattern.findall(csv_i)  # 在字符串 csv_i 中查找与正则表达式模式 csv_pattern 匹配的所有子串
        if csv_result != []:
            csv_file_name.append(csv_result)

    for csv_name in csv_file_name:
        '''遍历csv文件名'''
        Name = str(csv_pattern.findall(csv_name[0])[0]).rstrip('_DAPI_path_view.csv')

        # edge_csv = pd.read_csv(file_a + big_file_name + '/' + Name + '_new_edge.csv', header=None)  # 读取细胞连接信息

        edge_path = os.path.join(file_a, big_file_name, Name + '_new_edge.csv')
        if not os.path.exists(edge_path):
            print(f"文件不存在: {edge_path}，跳过。")
            continue
        edge_csv = pd.read_csv(edge_path, header=None)

        node_kind = file_b + big_file_name + '/' + Name + '_DAPI_path_view_各类细胞坐标.csv'
        # 使用pandas的read_excel函数读取Excel文件
        df = pd.read_excel(node_kind)
        # 将DataFrame转换为字典，其中第一列的值作为键，第三列的值作为值
        node_kind_dict = df.set_index('编号')['细胞类型'].to_dict()

        if panel == 1:
            edge_filename = 'E:/研一/Community-Detection/dataset/新连接信息/panel1/' + Name + '_DAPI_path_view.csv_各类细胞连接信息.csv'

            with (open(edge_filename, 'w', newline='', ) as csvfile):
                # 创建CSV写入器
                writer = csv.writer(csvfile)
                # 指定列标题
                writer.writerow(
                    ['CD4-CD4连接边', 'CD20-CD20连接边', 'CD38-CD38连接边', 'CD66B-CD66B连接边', 'FOXP3-FOXP3连接边',
                     'CD4-CD20连接边', 'CD4-CD38连接边', 'CD4-CD66B连接边', 'CD4-FOXP3连接边', 'CD20-CD38连接边',
                     'CD20-CD66B连接边', 'CD20-FOXP3连接边', 'CD38-CD66B连接边', 'CD38-FOXP3连接边',
                     'CD66B-FOXP3连接边'])
                for node in range(0, len(edge_csv)):
                    if node_kind_dict[edge_csv[0][node]] == 'CD4' and node_kind_dict[edge_csv[1][node]] == 'CD4':
                        writer.writerow([f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'CD20' and node_kind_dict[edge_csv[1][node]] == 'CD20':
                        writer.writerow(['', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'CD38' and node_kind_dict[edge_csv[1][node]] == 'CD38':
                        writer.writerow(['', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'CD66B' and node_kind_dict[edge_csv[1][node]] == 'CD66B':
                        writer.writerow(['', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'FOXP3' and node_kind_dict[edge_csv[1][node]] == 'FOXP3':
                        writer.writerow(['', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '',
                                         '', '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD4' and node_kind_dict[edge_csv[1][node]]
                        == 'CD20') or (node_kind_dict[edge_csv[0][node]] == 'CD20' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD4'):
                        writer.writerow(['', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '',
                                         '', '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD4' and node_kind_dict[edge_csv[1][node]]
                        == 'CD38') or (node_kind_dict[edge_csv[0][node]] == 'CD38' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD4'):
                        writer.writerow(['', '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}",
                                         '', '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD4' and node_kind_dict[edge_csv[1][node]]
                        == 'CD66B') or (node_kind_dict[edge_csv[0][node]] == 'CD66B' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD4'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD4' and node_kind_dict[edge_csv[1][node]]
                        == 'FOXP3') or (node_kind_dict[edge_csv[0][node]] == 'FOXP3' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD4'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD20' and node_kind_dict[edge_csv[1][node]]
                        == 'CD38') or (node_kind_dict[edge_csv[0][node]] == 'CD38' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD20'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD20' and node_kind_dict[edge_csv[1][node]]
                        == 'CD66B') or (node_kind_dict[edge_csv[0][node]] == 'CD66B' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD20'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD20' and node_kind_dict[edge_csv[1][node]]
                        == 'FOXP3') or (node_kind_dict[edge_csv[0][node]] == 'FOXP3' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD20'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD38' and node_kind_dict[edge_csv[1][node]]
                        == 'CD66B') or (node_kind_dict[edge_csv[0][node]] == 'CD66B' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD38'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD38' and node_kind_dict[edge_csv[1][node]]
                        == 'FOXP3') or (node_kind_dict[edge_csv[0][node]] == 'FOXP3' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD38'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD66B' and node_kind_dict[edge_csv[1][node]]
                        == 'FOXP3') or (node_kind_dict[edge_csv[0][node]] == 'FOXP3' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD66B'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}"])
        if panel == 2:
            # edge_filename = 'E:/研一/Community-Detection/dataset/result2/' + Name + '_DAPI_path_view.csv_各类细胞连接信息.csv'
            edge_filename = 'E:/研一/Community-Detection/dataset/新连接信息/panel2/' + Name + '_DAPI_path_view.csv_各类细胞连接信息.csv'

            with (open(edge_filename, 'w', newline='') as csvfile):
                # 创建CSV写入器
                writer = csv.writer(csvfile)
                # 指定列标题
                writer.writerow(
                    ['CD8-CD8连接边', 'CD68-CD68连接边', 'CD133-CD133连接边', 'CD163-CD163连接边', 'PDL1-PDL1连接边',
                     'CD8-CD68连接边', 'CD8-CD133连接边', 'CD8-CD163连接边', 'CD8-PDL1连接边', 'CD68-CD133连接边',
                     'CD68-CD163连接边', 'CD68-PDL1连接边', 'CD133-CD163连接边', 'CD133-PDL1连接边',
                     'CD163-PDL1连接边'])
                for node in range(0, len(edge_csv)):
                    if node_kind_dict[edge_csv[0][node]] == 'CD8' and node_kind_dict[edge_csv[1][node]] == 'CD8':
                        writer.writerow([f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'CD68' and node_kind_dict[edge_csv[1][node]] == 'CD68':
                        writer.writerow(['', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'CD133' and node_kind_dict[edge_csv[1][node]] == 'CD133':
                        writer.writerow(['', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'CD163' and node_kind_dict[edge_csv[1][node]] == 'CD163':
                        writer.writerow(['', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '',
                                         '', '', '', '', '', '', '', ''])
                    elif node_kind_dict[edge_csv[0][node]] == 'PDL1' and node_kind_dict[edge_csv[1][node]] == 'PDL1':
                        writer.writerow(['', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '',
                                         '', '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD8' and node_kind_dict[edge_csv[1][node]]
                        == 'CD68') or (node_kind_dict[edge_csv[0][node]] == 'CD68' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD8'):
                        writer.writerow(['', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '',
                                         '', '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD8' and node_kind_dict[edge_csv[1][node]]
                        == 'CD133') or (node_kind_dict[edge_csv[0][node]] == 'CD133' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD8'):
                        writer.writerow(['', '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}",
                                         '', '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD8' and node_kind_dict[edge_csv[1][node]]
                        == 'CD163') or (node_kind_dict[edge_csv[0][node]] == 'CD163' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD8'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD8' and node_kind_dict[edge_csv[1][node]]
                        == 'PDL1') or (node_kind_dict[edge_csv[0][node]] == 'PDL1' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD8'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD68' and node_kind_dict[edge_csv[1][node]]
                        == 'CD133') or (node_kind_dict[edge_csv[0][node]] == 'CD133' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD68'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD68' and node_kind_dict[edge_csv[1][node]]
                        == 'CD163') or (node_kind_dict[edge_csv[0][node]] == 'CD163' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD68'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD68' and node_kind_dict[edge_csv[1][node]]
                        == 'PDL1') or (node_kind_dict[edge_csv[0][node]] == 'PDL1' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD68'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD133' and node_kind_dict[edge_csv[1][node]]
                        == 'CD163') or (node_kind_dict[edge_csv[0][node]] == 'CD163' and
                                        node_kind_dict[edge_csv[1][node]] == 'CD133'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", '', ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD133' and node_kind_dict[edge_csv[1][node]]
                        == 'PDL1') or (node_kind_dict[edge_csv[0][node]] == 'PDL1' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD133'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}", ''])
                    if (node_kind_dict[edge_csv[0][node]] == 'CD163' and node_kind_dict[edge_csv[1][node]]
                        == 'PDL1') or (node_kind_dict[edge_csv[0][node]] == 'PDL1' and
                                       node_kind_dict[edge_csv[1][node]] == 'CD163'):
                        writer.writerow(['', '', '', '', '', '', '',
                                         '', '', '', '', '', '', '', f"{edge_csv[0][node]};{edge_csv[1][node]}"])
