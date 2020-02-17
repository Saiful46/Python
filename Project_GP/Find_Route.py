import pandas as pd

f_n = 'data2_moderate'
out_path = f_n + '_out.xlsx'
df = pd.read_excel(f_n + '.xlsx', sheet_name=1, skiprows=0)
edge = df['EDGE'].tolist()
s = df['Source'].tolist()
d = df['Destination'].tolist()


def join_string(list_string):
    # Join the string based on '-' delimiter
    string = '>'.join(list_string)

    return string

def find_all_paths(graph, start, end, path=[]):
    path = path + [start]
    if start == end:
        return [path]
    if start not in graph:
        return []
    paths = []
    for node in graph[start]:
        if node not in path:
            newpaths = find_all_paths(graph, node, end, path)
            for newpath in newpaths:
                paths.append(newpath)
    #print(paths)
    return paths

if __name__ == "__main__":

    e_list = []

    for i in range(len(edge)):
        print(edge[i])
        e = edge[i].split('-')
        e_list.append([e[0], e[1]])

    graph = {}
    for key, val in e_list:
        graph.setdefault(key, []).append(val)

    src_dst_route = []
    src_dst = []

    header = ['Source', 'Destination']

    route_length = []
    for i in range(len(s)):
        g = find_all_paths(graph, s[i], d[i])
        src_dst.append([s[i] , d[i]])

        route = []
        for j in range(len(g)):
            new_string = join_string(g[j])
            route.append(new_string)
        route_length.append(len(route))
        src_dst_route.append(src_dst[i] + route)

    header.extend(['Route' for i in range(max(route_length))])
    print(header)

    # ---------------------write data file using panda-----------------------------------
    df1 = pd.DataFrame(src_dst_route, columns = header)

    writer = pd.ExcelWriter('data2_route_out2.xlsx', engine='xlsxwriter')

    df1.to_excel(writer, sheet_name='Sheet1', index=False)  # for sheet1 append values

    writer.save()










