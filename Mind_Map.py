import graphviz
import openpyxl

file_name = "your_file.xlsx"
wb = openpyxl.load_workbook(file_name)
sheet = wb.active

data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    data.append(row)

new_file_name = str(data[0][0])

graph = graphviz.Digraph(format='png', engine='dot', graph_attr={'rankdir': 'TB', 'size': '20,20'})

added_edges = set()

root = str(data[0][0])
graph.node(root, shape='box', style='filled', fillcolor='lightblue', fontsize='12')

for row in data:
    non_empty_cells = [str(cell) for cell in row if cell]
    for i in range(len(non_empty_cells) - 1):
        parent = non_empty_cells[i]
        child = non_empty_cells[i + 1]
        edge = (parent, child)
        if edge not in added_edges:
            added_edges.add(edge)
            graph.node(child, shape='box', style='filled', fillcolor='lightblue', fontsize='12')
            graph.edge(parent, child, constraint='true', fontsize='10')

output_file = new_file_name
graph.render(output_file, cleanup=True)
