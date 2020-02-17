import pandas as pd
from collections import Counter


f_n = input("Enter the file name : ")
#f_n = 'data2_moderate'
out_path = f_n+'_out.xlsx'
df = pd.read_excel(f_n+'.xlsx', sheet_name=0, skiprows=0)

node = df['Source'].tolist()
route = df['Route'].tolist()

node_Route_edge = []   #for first sheet output
edge_Length_Nodes = [] #for second sheet output
nodeU_nodeC = []       #for third sheet output
edge = []
edge_node = []
header_1 = ['Source', 'Route']

l_e = []
node_in_route = []
for i in range(len(df)):
    e = route[i].split('>')
    node_in_route.extend(e)
    l_e.append(len(e)-1)

    for l in range(len(e)-1):
        edge.append(e[l]+"-"+e[l+1])
        edge_node.append([(e[l]+"-"+e[l+1]), node[i]])

    node_Route_edge.append(([node[i] , route[i]]) + edge)
    edge.clear()

print("node in route : ", node_in_route)

header_1.extend(['Edge' for i in range(max(l_e))])

#---------converting edge_node list to edge_node Dictionary------------------
edge_node_Dict = {}
for key, val in edge_node:
    edge_node_Dict.setdefault(key, []).append(val)

#---------extracting set of edge + no. of nodes and values of nodes----------------
#edge_Length_Nodes.append(['EDGE', 'NO. OF NODES', "<<", '---', '---', '---', "NODES",'---','---','---', '>>'])
header_2 = ['Edge', 'No. Of Nodes']
l_n = []
for key, value in edge_node_Dict.items():
    nodes_list = ([item for item in value if item])
    nodes_list_length = len([item for item in value if item])
    l_n.append(nodes_list_length)
    edge_Length_Nodes.append([key] + [nodes_list_length] + nodes_list)  # appending the Second file output


header_2.extend(['Nodes' for i in range(max(l_n))])

#print(header_2)

#---------------------find no of same node in route---------------------------------
header_3 = ['Unique Node', 'Count']
U_node = list(Counter(node_in_route).keys())
C_node = list(Counter(node_in_route).values())

for i in range(len(U_node)):
    nodeU_nodeC.append( [ U_node[i], C_node[i] ] )


#---------------------write data file using panda-----------------------------------
df1 = pd.DataFrame(node_Route_edge, columns = header_1)
df2 = pd.DataFrame(edge_Length_Nodes, columns = header_2)
df3 = pd.DataFrame(nodeU_nodeC, columns = header_3)

writer = pd.ExcelWriter(out_path, engine='xlsxwriter')

df1.to_excel(writer, sheet_name='Sheet1', index = False) # for sheet1 append values
df2.to_excel(writer, sheet_name='Sheet2', index = False) # for sheet2 append values
df3.to_excel(writer, sheet_name='Sheet3', index = False) # for sheet3 append values

writer.save()