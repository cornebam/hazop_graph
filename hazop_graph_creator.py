from openpyxl import load_workbook
import gmlcreator_custom as gmlc
import networkx as nx
import argparse

RISK_MAPPING = {"A": {"id": 6, "colour": "#FF0000"},
                "B": {"id": 5, "colour": "#FF9900"},
                "C": {"id": 4, "colour": "#FFFF00"},
                "D": {"id": 3, "colour": "#FF00FF"},
                "E": {"id": 2, "colour": "#00FF00"},
                "F": {"id": 1, "colour": "#CCFFCC"}
               }

class IdGenerator():
    def __init__(self):
        self.id = -1

    def g(self):
        self.id = self.id + 1
        return self.id

class GraphCustom(nx.DiGraph):
    def __init__(self, rows):
        super().__init__()
        self.create_graph_from_rows(rows)

    def create_graph_from_rows(self, rows):
        i = IdGenerator()
        for row in rows:
            # Updating Groups and affiliated nodes from the first section
            cur_group = row[3].strip()
            cur_node = (row[5] + " " + row[6]).strip().lower()
            if cur_group not in nx.get_node_attributes(self, "label").values():
                self.add_group(cur_group, i.g())
            cur_group_id = self.get_id_from_name(cur_group)
            if self.node_not_exists_in_group(cur_node, cur_group_id):
                cur_node_id = i.g()
                self.add_node(cur_node_id, label=cur_node, gid=cur_group_id, 
                               id=cur_node_id, graphics={"type": "ellipse"})
    
            # Updating Groups and affiliated nodes from the cause section
            cause_group = row[10].strip()
            cause_node = row[8].strip().lower()
            if cause_group not in nx.get_node_attributes(self, "label").values():
                self.add_group(cause_group, i.g())
            cause_group_id = self.get_id_from_name(cause_group)
            if self.node_not_exists_in_group(cause_node, cause_group_id):
                cause_node_id = i.g()
                self.add_node(cause_node_id, label=cause_node, gid=cause_group_id,
                               id=cause_node_id, graphics={"type": "ellipse"})
    
            # Updating Groups and affiliated nodes from the implication
            implic_group = row[14].strip()
            implic_node = row[12].strip().lower()
            if implic_group not in nx.get_node_attributes(self, "label").values():
                self.add_group(implic_group, i.g())
            implic_group_id = self.get_id_from_name(implic_group)
            if self.node_not_exists_in_group(implic_node, implic_group_id):
                implic_node_id = i.g()
                self.add_node(implic_node_id, label=implic_node, id=implic_node_id,
                               gid=implic_group_id, graphics={"type": "ellipse"})
            
            # Adding edges
            cur_node_id = self.get_id_from_name(cur_node, cur_group)
            cause_node_id = self.get_id_from_name(cause_node, cause_group)
            implic_node_id = self.get_id_from_name(implic_node, implic_group)
            if not cause_node_id == cur_node_id:
                self.add_edge(cause_node_id, cur_node_id)
            if not cur_node_id == implic_node_id:
                self.add_edge(cur_node_id, implic_node_id)            
                
            #Adding risk to implicit nodes and 
            node_risk = RISK_MAPPING[row[18]]["id"]
            self.update_node_risk(implic_node_id, node_risk)
        
    def add_group(self, name, group_id):
        self.add_node(group_id, label=name, id=group_id, isGroup=1)

    def node_not_exists_in_group(self, node, group_id):
        for node_id in self.nodes():
            if self.node[node_id]["label"] == node:
                if self.node[node_id]["gid"] == group_id:
                    return False
        return True

    def get_id_from_name(self, name, group=None):
        group_id = None
        if group:
            group_id = self.get_id_from_name(group)
            for node_id in self.nodes():
                label = self.node[node_id]["label"]
                if self.node[node_id]["label"] == name:
                    if self.node[node_id]["gid"] == group_id:
                        return self.node[node_id]["id"]
        else:
            for node_id in self.nodes():
                if self.node[node_id]["label"] == name:
                    return self.node[node_id]["id"]
    
    def update_node_risk(self, node_id, node_risk):
        if ((not self.node[node_id].get("risk") or 
             self.node[node_id]["risk"] < node_risk)):
            self.node[node_id]["risk"] = node_risk

    def set_backdated_risks(self):
        for node_id in nx.get_node_attributes(self, "risk"):
            node_risk = self.node[node_id]["risk"]
            predecessors = self.get_all_predecessors(node_id)
            print(predecessors)
            if predecessors:
                for predecessor in predecessors:
                    self.update_node_risk(predecessor, node_risk)
            
    def get_all_predecessors(self, node_id):
        global_predecessors = set()
        nodes_to_add = set([node_id])
        while nodes_to_add:
            global_predecessors.update(nodes_to_add)
            local_predecessors = set()
            for node in nodes_to_add:
                local_predecessors.update([i for i in self.predecessors(node)])
            nodes_to_add = local_predecessors.difference(global_predecessors)
        return global_predecessors

    def get_all_successors(self, node_id):
        global_successors = set()
        nodes_to_add = set([node_id])
        while nodes_to_add:
            global_successors.update(nodes_to_add)
            local_successors = set()
            for node in nodes_to_add:
                local_successors.update([i for i in self.successors(node)])
            nodes_to_add = local_successors.difference(global_successors)
        return global_successors
    
    def get_single_node(self, node_name, group_name):
        node_id = self.get_id_from_name(node_name, group_name)
        print(node_id)
        predecessors = self.get_all_predecessors(node_id)
        successors = self.get_all_successors(node_id)
        relatives = predecessors.union(successors)
        nodes_to_remove = set()
        for node in self.nodes():
            if node not in relatives and not self.node[node].get("isGroup"):
                nodes_to_remove.add(node)
        for node in nodes_to_remove:
            self.remove_node(node)

    def colour_nodes(self):
        for node_id in self.nodes():
            node_risk = self.node[node_id].get("risk")
            if node_risk:
                self.node[node_id]["graphics"]["fill"] = risk_colour(node_risk)

    def colour_edges(self):
        for node_id in self.nodes():
            node_risk = self.node[node_id].get("risk")
            if node_risk:
                for edge in self.edges():
                    source = edge[0]
                    target = edge[1]
                    if node_id in [source, target]:
                        edge_graphics = {"width": node_risk, 
                                         "fill": risk_colour(node_risk),
                                         "targetArrow":	"standard"}
                        self.edges[source,target]["graphics"] = edge_graphics
    
    def limit_risk_range(self, risk):
        risk_id = RISK_MAPPING[risk]["id"]
        nodes_to_remove = set()
        for node_id in self.nodes():
            try:
                if self.node[node_id]["risk"] < risk_id:
                    nodes_to_remove.add(node_id)
            except KeyError:
                continue
        for node in nodes_to_remove:
            self.remove_node(node)
    
    def remove_remaining_edges(self):
        nodes = self.nodes()
        for edge in self.edges():
            source = edge[0]
            target = edge[1]
            if source not in nodes or target not in nodes:
                self.remove_edge(source, target)

def get_xls_rows(file_, tab):
    wb = load_workbook(filename = file_)
    ws = wb[tab]
    rows = []
    for k, row in enumerate(ws):
        if k < 4: continue
        current_row = []
        for cell in row:
            current_row.append(cell.value)
        if current_row[7] == 'not relevant':
            continue
        rows.append(current_row)
    return rows

def risk_colour(risk_id):
    for risk in RISK_MAPPING.values():
        if risk_id in risk.values():
            return risk["colour"]

def main(args):
    
    rows = get_xls_rows(args.file, args.tab)
    graph = GraphCustom(rows)
    graph.create_graph_from_rows(rows)

    
    print(args.single_node)
    if args.single_node:
        node = args.single_node[0].replace("_", " ")
        group = args.single_node[1].replace("_", " ")
        graph.get_single_node(node, group)
    #Apply risk to predescessor nodes
    graph.set_backdated_risks()
    
    graph.limit_risk_range(args.risk)
    if args.colour_nodes: graph.colour_nodes()
    if args.colour_edges: graph.colour_edges()
    graph.remove_remaining_edges()
    gmlc.write_gml(graph, args.output)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-o', '--output', type=str, default="gmlcreator_test.gml",
                        help='path to target gmlfile')
    parser.add_argument('-r', '--risk', default="F",
                        help='minimum risk level to show (F to A)')
    parser.add_argument('--colour-nodes', default=True, type=bool,
                        help='improves visualization of higher risks')
    parser.add_argument('--colour-edges', default=True, type=bool,
                        help='improves visualization of higher risks ')
    parser.add_argument('-s', '--single-node', nargs="+", type=str,
                        help='Choose a single node and all its successors \
                        and predecessors. Syntax: -s no_flow C2')
    parser.add_argument('-f', '--file', required=False, type=str, 
                        default='HAZOP-Beispiel.xlsm',
                        help='input hazop xls/xlsm file')
    parser.add_argument('-t', '--tab', required=False, default='mHAZOP - Module',
                        help='name of the tab to use')

    args = parser.parse_args()
    main(args)
