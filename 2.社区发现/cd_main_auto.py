import pandas as pd
import csv
import time
from cdlib import algorithms
from cdlib import NodeClustering
import networkx as nx
import os
from networkx.algorithms.community import greedy_modularity_communities
from networkx.algorithms.community import naive_greedy_modularity_communities
# from networkx.algorithms.community import lukes_partitioning
from networkx.algorithms.community import asyn_lpa_communities
from networkx.algorithms.community import label_propagation_communities
from networkx.algorithms.community import asyn_fluidc
from networkx.algorithms.community import k_clique_communities
from networkx.algorithms.community import girvan_newman
import math
import networkx.algorithms.community as nx_comm
from collections import Counter
from scipy.spatial import distance
import re
from collections import defaultdict

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


file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 1/P2017004-P117-2-T5-1-20200217/'
# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 1/P2017004-P117-2-T5-2-20200217/'
# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 2/117-3-T5-1/'
# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 2/117-3-T5-2/'
# file_a = 'E:/研一/Community-Detection/dataset/20200605-M-2019153LC-The First Affiliated Hospital of Guangzhou Medical University - Liang Wenhua/Panel 2/117-3-T5-3/'

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

    # 得到文件夹下的CSV文件的路径
    for csv_i in csv_l:
        csv_result = csv_pattern.findall(csv_i)  # 在字符串 csv_i 中查找与正则表达式模式 csv_pattern 匹配的所有子串
        if csv_result != []:
            csv_file_name.append(csv_result)

    
    for csv_name in csv_file_name:
        '''遍历csv文件名'''
        # df_csv = pd.read_csv(file_a + big_file_name + '/' + big_file_name + '/' + str(csv_name[0]), usecols=[1])

        Name = str(csv_pattern.findall(csv_name[0])[0]).rstrip('_DAPI_path_view.csv')

        edge_filename = file_a + big_file_name + '/' + Name + '_edge.csv'
        edge_csv = pd.read_excel(file_a + big_file_name + '/' + Name + '_边表.csv', usecols=[0])  # 读取边表文件的第一列
        data_1 = edge_csv['细胞的边'].str.split(',', expand=True)  # 选择名为'细胞的边'的列
        data_2 = data_1[0].str.split('(', expand=True)
        data_3 = data_1[1].str.split(')', expand=True)
        edge_l = data_2.shape[0]  # 储存data_2的行数

        # 生成文件名为edge的文件
        with open(edge_filename, 'w', encoding='utf8') as fe:
            for i in range(0, edge_l):
                fe.writelines(data_2[1][i])
                fe.writelines(',')
                fe.writelines(data_3[0][i])
                fe.writelines('\n')
            fe.close()

        degree_filename = file_a + big_file_name + '/' + Name + '_degree.csv'

        # 读取边表CSV文件并计算度
        def calculate_degrees(edge_filename):  
            degr = defaultdict(int)  # 初始化一个字典来存储每个顶点的度
            with open(edge_filename, 'r', newline='') as csvfile:  # 打开边表CSV文件进行读取
                reader = csv.reader(csvfile)
                for row in reader:
                    vertex1, vertex2 = row  # 假设每行包含两个顶点，表示一条边
                    # 增加顶点的度
                    degr[vertex1] += 1
                    degr[vertex2] += 1
            return degr

        def sort_vertices_by_degree(degrees):
            return sorted(degrees.items(), key=lambda item: item[1], reverse=True)

        def write_degrees_to_csv(sorted_degrees, degree_filename):  # 将度信息写入到新的CSV文件
            with open(degree_filename, 'w', newline='') as csvfile:  # 打开新的CSV文件进行写入
                writer = csv.writer(csvfile)
                # writer.writerow(['Vertex', 'Degree'])  # 写入标题行
                for vertex, degree in sorted_degrees:  # 写入度信息
                    writer.writerow([vertex, degree])

        degrees = calculate_degrees(edge_filename)  # 计算度
        sorted_degrees = sort_vertices_by_degree(degrees)  # 按度降序排序顶点
        write_degrees_to_csv(sorted_degrees, degree_filename)  # 将度写入到新的CSV文件


        dis_filename = file_a + big_file_name + '/' + Name + '_dis.csv'
        dis_csv = pd.read_excel(file_a + big_file_name + '/' + Name + '_各类细胞坐标.csv')
        data_5 = dis_csv['细胞坐标'].str.split(',', expand=True)
        data_6 = data_5[0].str.split('(', expand=True)  # x坐标的矩阵
        data_7 = data_5[1].str.split(')', expand=True)  # y坐标的矩阵
        data_8 = dis_csv['编号']
        dis_l = data_7.shape[0]  # 储存data_7的行数
        with open(dis_filename, 'w', encoding='utf8') as fd:
            for i in range(0, dis_l):
                fd.writelines(str(i))
                fd.writelines(',')
                fd.writelines(data_6[1][i])
                fd.writelines(',')
                fd.writelines(data_7[0][i])
                fd.writelines('\n')
            fd.close()


        class nbcd_community(dict):  # 定义类

            # __init__ function
            def __init__(self):
                self = dict()

            # Function to add key:value 键值对
            def add(self, key, value):
                self[key] = value


        dict_obj = nbcd_community()  # 创建类的实列（节点所属社区的键值对）
        degree_obj = nbcd_community()  # 创建类的实列（节点与度的键值对）
        dis_obj = nbcd_community()  # 创建类的实列（节点之间距离的键值对）
        # print("enter the value of alpha: alpha value must be '1.5' or '2' or '2.5' or '3'")
        # alpha = float(input())
        # alpha = 1.5
        # alpha = 2
        alpha = 2.5
        # alpha = 3
        G = nx.Graph()  # 创建一个新的空无向图

        fpath = str(file_a + big_file_name + '/' + Name + '_edge.csv')
        G = nx.read_edgelist(fpath, delimiter=',')  # ","为分隔符,从文件中读取边列表
        community = []
        c1 = []
        community.append(c1)
        community_count = 0
        visited = []
        not_visited = []

        deg_path = str(file_a + big_file_name + '/' + Name + '_degree.csv')
        with open(deg_path, 'r', encoding='utf-8-sig') as t1:
            fileone = t1.readlines()
        for line in fileone:
            x = line.split(',', 1)[0]  # 通过","分隔符进行切片，分割成两个，取第一个元素赋值给x
            y = int(line.split(',', 1)[-1])  # 取第二个元素赋值给y
            degree_obj.add(x, y)  # 在集合里添加元素，x是键，y是值
            not_visited.append(x)  # 在列表结尾添加一元素
            visited.append(x)

        dis_path = str(file_a + big_file_name + '/' + Name + '_dis.csv')
        with open(dis_path, 'r', encoding='utf-8-sig') as td:  
            filedis = td.readlines()
        for line in filedis:
            x = line.split(',', 2)[0]  # 通过","分隔符进行切片，分割成两个，取第一个元素赋值给x
            xd = int(line.split(',', 2)[1])  # 取第二个元素赋值给xd
            yd = int(line.split(',', 2)[2])  # 取第二个元素赋值给yd
            dis_cor = (xd, yd)  # 坐标
            dis_obj.add(x, dis_cor)  # 在集合里添加元素，x是键，dis_cor是值

        start_time = time.time()
        ################################################################START#########################################################################
        x = not_visited[0]  # 将第一个元素赋值给x,即取目前度最大的那个节点
        deg_x = degree_obj[x]  # 节点x的度,x为键，这里取了它的值
        dis_x = dis_obj[x]  # 节点x的坐标位置
        community[community_count].append(x)  # 把x添加到社区中
        not_visited.remove(x)  # 从未访问节点列表中移除某个节点x
        nbr_of_x = list(G[x])  # 返回图的邻居节点列表(获取所有邻居)
        dict_obj.add(x, community_count)  # 添加节点x与所属社区键值对

        while len(not_visited) != 0:
            x = not_visited[0]
            deg_x = degree_obj[x]
            check = 0  # sim（x，c）的值
            if deg_x != 2 and deg_x != 1:
                deg_x = degree_obj[x]
                nbr_x = list(G[x])  # G[x] 和 list(G[x]) 在功能上通常是等效的，但后者确保返回一个列表。
                for nod in nbr_x:
                    try:
                        n = dict_obj(nod)
                        count = len(set(nbr_x).intersection(set(community[n])))
                        # 返回x的邻居节点与社区c中节点的交集的长度，即x的邻居节点属于社区c的个数
                        if count > deg_x / alpha:
                            community[n].append(x)
                            not_visited.remove(x)
                            dict_obj.add(x, n)
                            check = 1
                            break
                    except Exception:
                        pass  # 忽略异常

                if check == 1:
                    continue
            if check != 1 and deg_x != 2 and deg_x != 1:  # otherwise
                new_community = []
                community.append(new_community)
                community_count = community_count + 1
                community[community_count].append(x)
                not_visited.remove(x)
                nbr_of_x = G[x]

                dict_obj.add(x, community_count)  # 把x归入到新创建的社区
                for y in nbr_of_x:  # 计算sim（x，y）
                    nbr_y = list(G[y])
                    deg_y = degree_obj[y]
                    dis_y = dis_obj[y]  # 节点y的坐标位置
                    for yn in nbr_y:
                        nbr_dis_y = dis_obj[yn]
                        dis = distance.euclidean(dis_x, dis_y)  # x与y两节点间的距离
                        dis_nbr = distance.euclidean(dis_y, nbr_dis_y)  # y与y除了x的另一个邻居节点间的距离
                        count = 0
                        if deg_y == 1 and y in not_visited:  # y的度为1（即y只与x相连），并且y还没有被访问过
                            community[community_count].append(y)
                            not_visited.remove(y)
                            dict_obj.add(y, community_count)
                        elif deg_y == 2 and dis < dis_nbr and y in not_visited:
                            # y的度为2（即y与x以及另一个邻居节点相连），并且y还没有被访问过,并且x与y的距离小于y与另一个邻居节点的距离
                            community[community_count].append(y)
                            not_visited.remove(y)
                            dict_obj.add(y, community_count)
                        else:
                            count = len(set(nbr_y).intersection(set(nbr_of_x))) + 1
                            if count > deg_y / alpha and y in not_visited:
                                community[community_count].append(y)
                                not_visited.remove(y)
                                dict_obj.add(y, community_count)

                continue
            if check != 1 and deg_x == 2:  # x只有两个邻居
                nbr_x = list(G[x])
                nbr_1 = nbr_x[0]  # 邻居1
                nbr_2 = nbr_x[1]  # 邻居2
                nbr_dis_1 = dis_obj[nbr_1]
                nbr_dis_2 = dis_obj[nbr_2]
                dis_nbr1 = distance.euclidean(dis_x, nbr_dis_1)
                dis_nbr2 = distance.euclidean(dis_x, nbr_dis_2)
                try:
                    comm_nbr1 = dict_obj[nbr_1]
                except Exception:
                    comm_nbr1 = -1  # 邻居1没有被访问过
                try:
                    comm_nbr2 = dict_obj[nbr_2]
                except Exception:
                    comm_nbr2 = -1  # 邻居2没有被访问过
                deg_nbr_1 = degree_obj[nbr_1]
                deg_nbr_2 = degree_obj[nbr_2]
                if (
                        deg_nbr_1 > deg_nbr_2 or dis_nbr1 < dis_nbr2) and comm_nbr1 != -1:
                    # 邻居1的度（距离）大于邻居2的度（距离），并且邻居1有被访问过
                    community[comm_nbr1].append(x)
                    not_visited.remove(x)
                    dict_obj.add(x, comm_nbr1)  # x直接归入邻居1所在的社区中
                elif (
                        deg_nbr_1 > deg_nbr_2 or dis_nbr1 < dis_nbr2) and comm_nbr1 == -1:
                    # 邻居1的度（距离）大于邻居2的度（距离），并且邻居1没有被访问过
                    new_community = []
                    new_community.append(x)
                    new_community.append(nbr_1)
                    community.append(new_community)
                    community_count = community_count + 1
                    dict_obj.add(x, community_count)
                    dict_obj.add(nbr_1, community_count)
                    not_visited.remove(x)
                    not_visited.remove(nbr_1)  # 创建一个新的社区，把x和邻居1都归入到其中
                elif (
                        deg_nbr_2 > deg_nbr_1 or dis_nbr1 > dis_nbr2) and comm_nbr2 != -1:
                    # 邻居2的度（距离）大于邻居1的度（距离），并且邻居2有被访问过
                    community[comm_nbr2].append(x)
                    not_visited.remove(x)
                    dict_obj.add(x, comm_nbr2)  # x直接归入邻居2所在的社区中
                elif (
                        deg_nbr_2 > deg_nbr_1 or dis_nbr1 > dis_nbr2) and comm_nbr2 == -1:
                    # 邻居2的度（距离）大于邻居1的度（距离），并且邻居2没有被访问过
                    new_community = []
                    new_community.append(x)
                    new_community.append(nbr_2)
                    community.append(new_community)
                    community_count = community_count + 1
                    dict_obj.add(x, community_count)
                    dict_obj.add(nbr_2, community_count)
                    not_visited.remove(x)
                    not_visited.remove(nbr_2)  # 创建一个新的社区，把x和邻居2都归入到其中
                elif (
                        deg_nbr_1 == deg_nbr_2 or dis_nbr1 == dis_nbr2) and comm_nbr1 != -1:
                    # 邻居1的度（距离）等于邻居2的度（距离），并且邻居1有被访问过
                    community[comm_nbr1].append(x)
                    not_visited.remove(x)
                    dict_obj.add(x, comm_nbr1)  # x直接归入邻居1所在的社区中
                elif (
                        deg_nbr_1 == deg_nbr_2 or dis_nbr1 == dis_nbr2) and comm_nbr2 != -1:
                    # 邻居1的度（距离）等于邻居2的度（距离），并且邻居2有被访问过
                    community[comm_nbr2].append(x)
                    not_visited.remove(x)
                    dict_obj.add(x, comm_nbr2)  # x直接归入邻居2所在的社区中
                else:  # 创建一个新的社区，把x和邻居1，邻居2一起归入其中
                    new_community = []
                    new_community.append(x)
                    new_community.append(nbr_1)
                    new_community.append(nbr_2)
                    community.append(new_community)
                    community_count = community_count + 1
                    dict_obj.add(x, community_count)
                    dict_obj.add(nbr_1, community_count)
                    dict_obj.add(nbr_2, community_count)
                    not_visited.remove(x)
                    not_visited.remove(nbr_1)
                    not_visited.remove(nbr_2)
                continue
            if deg_x == 1:  # x只有一个邻居节点，如果邻居节点被访问过，则把x归入邻居节点所在社区。否则创建一个新社区，把这两点归入其中
                nbr_x = list(G[x])
                try:
                    t = dict_obj[nbr_x[0]]
                    community[t].append(x)
                    not_visited.remove(x)
                    check = 1
                    dict_obj.add(x, t)
                except Exception:
                    pass
                if check != 1:
                    new_community = []
                    community.append(new_community)
                    community_count = community_count + 1
                    community[community_count].append(x)
                    dict_obj.add(x, community_count)
                    for u in nbr_x:
                        community[community_count].append(u)
                        if u in not_visited:
                            not_visited.remove(u)
                            dict_obj.add(u, community_count)
                    not_visited.remove(x)
                    check = 1
        # -----------------------------------------------Shifting Nodes-------------------------------------------------------------
        shift = 1  # 初始化转移判断条件
        mod_old = nx_comm.modularity(G, community)  # 计算模块度，模块度越高，社区划分得越好
        while shift == 1:
            check = 0
            for a in G.nodes:
                nbr_a = G[a]  # 获取节点a的邻居
                community_id = []
                current_nbr_in_comm_x = 0
                for n in nbr_a:  # 遍历邻居节点集合中的每个节点n
                    if dict_obj[n] == dict_obj[a]:  # 检查节点n与节点a是否在同一个社区
                        current_nbr_in_comm_x = current_nbr_in_comm_x + 1
                if current_nbr_in_comm_x > degree_obj[a] / alpha:
                    continue  # 跳过当前节点的处理
                for b in nbr_a:
                    community_id.append(dict_obj[b])
                df = pd.DataFrame({'Number': community_id})
                # 使用这些邻居的标识符创建一个DataFrame，并计算每个标识符的出现次数
                df1 = pd.DataFrame(data=df['Number'].value_counts())
                df1['Count'] = df1['Number'].index  # index为行标签
                df1 = list(df1[df1['Number'] == df1.Number.max()]['Count'])
                max_comm = df1[0]  # 找到出现次数最多的标识符，并将其赋值给max_comm
                max_degree = 0
                if len(df1) != 1:
                    # 如果存在多个出现次数最多的标识符（即不止一个），则选择具有最大度数的节点所属的社区作为新的社区标识符
                    for t in df1:
                        deg_count = 0
                        for b in nbr_a:
                            if dict_obj[b] == t:
                                deg_count = deg_count + degree_obj[b]
                        if deg_count > max_degree:
                            max_degree = deg_count
                            max_comm = t
                    if dict_obj[a] != max_comm:
                        # 如果节点的当前社区标识符与新找到的社区标识符不同，则将该节点从当前社区中移除，并添加到新的社区中
                        community[dict_obj[a]].remove(a)
                        dict_obj[a] = max_comm
                        community[dict_obj[a]].append(a)
                        check = 1
                elif dict_obj[a] != df1[0]:
                    community[dict_obj[a]].remove(a)
                    dict_obj[a] = df1[0]
                    community[dict_obj[a]].append(a)
                    check = 1  # 表示社区归属发生了更改
            modu = nx_comm.modularity(G, community)
            if modu - mod_old > 0.01:
                print("Old Modularity = ", mod_old, ", New Modularity = ", modu)
                print("New Modularity is greater than old Modularity, so one more round of shifting will be done")
                mod_old = modu
                shift = 1
            else:
                print("Old Modularity = ", mod_old, ", New Modularity = ", modu)
                print(
                    "Old Modularity is greater than or approximately equal to new Modularity, so no more shifting round, Stop")
                shift = 0
        print("\nNBCD took", round(time.time() - start_time, 4), "sec to run")  # round()函数用于对浮点数进行四舍五入操作（4位小数）
        i = 0
        while i < len(community):
            if len(community[i]) == 0:
                community.remove(community[i])
            else:
                i = i + 1
        print("Number of communities = ", len(community))
        print("Modularity =", nx_comm.modularity(G, community))

#        print("\nDo you want to print Communities identified by NBCD? Type 'y' for yes and type 'n' for no.")
#        opt = str(input())
#        if opt == 'y':
#            print(community)
#        print("\nDo you want to write the communities into a file? Type 'y' for yes and type 'n' for no.")
#        opt1 = str(input())
#        if opt1 == 'y':
#        print("\nPlz enter the output file path")
#        opath = str(input())
        out_path = str(file_a + big_file_name + '/' + Name + '_DAPI_path_view_社区.csv')
        opath = out_path
#        opath1 = '1' + opath
        opath1 = str(file_a + big_file_name + '/' + Name + '_DAPI_path_view_社区(transfer).csv')
        with open(opath1, 'w') as outFile:
            for a in community:
                outFile.write(str(a) + '\n')
        with open(opath1, 'r') as fin, open(opath, 'w') as fout:
            data = fin.read()
            data = data.replace("'", "")
            data = data.replace("[", "")
            data = data.replace("]", "")

            fout.write(data)


        

#        print("\nFile successfully writen")

# ########################################################################END###################################################################
