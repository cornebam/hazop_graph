#!/usr/bin/env python

__author__ = "Cornelius Kirchhoff"
__email__ = "cornelius.kirchhoff@gmail.com"

from openpyxl import load_workbook
import gmlcreator_custom as gmlc
import networkx as nx
import argparse

RISK_MAPPING = {"A": {"id": 6, "colour": "#FA0018"},
                "B": {"id": 5, "colour": "#F4B98C"},
                "C": {"id": 4, "colour": "#FFCC66"},
                "D": {"id": 3, "colour": "#C2DE68"},
                "E": {"id": 2, "colour": "#7A9A40"},
                "F": {"id": 1, "colour": "#008000"}
               }
               
ROW_INDEXES = {"cur_group"   : 3,
               "cur_node"    : 5,
               "cause_group" : 10,
               "cause_node"  : 8,
               "implic_group": 14,
               "implic_node" : 12,
               "risk"        : 18
               }

class IdGenerator(object):
    """ This is just for returning an id for nodes and groups. """
    def __init__(self):
        self.id = -1

    def g(self): # get id
        self.id = self.id + 1
        return self.id

class HazopGraph(nx.DiGraph):
    def __init__(self, rows):
        super().__init__()
        self.rows = rows

    def create_graph_from_rows(self):
        """
        Read data from rows and fill graph with groups, nodes and edges.
        For that it goes through the defined indexes of table elements
        and reads the component name and the error.
        Note: Group = Component; Node = Error; Edge = Dependency between errors.
        For each Group-Node pair it checks if the group exists and creates
        it if it does not. Then it checks if the node exists in the group
        with the given name/label and creates it if it does not.
        Then the edges are added from cause to cur node and from cur to
        implic node. In the end the risk for the implic_node is set.
        """
        i = IdGenerator() # i.g() returns integer which ascends every time it is called
        rows = self.rows
        # Mapping for the indexes for easier adapting to other tables
        x1 = ROW_INDEXES["cur_group"]
        x2 = ROW_INDEXES["cur_node"]
        x3 = ROW_INDEXES["cause_group"]
        x4 = ROW_INDEXES["cause_node"]
        x5 = ROW_INDEXES["implic_group"]
        x6 = ROW_INDEXES["implic_node"]
        x7 = ROW_INDEXES["risk"]
        for row in rows:
            # Sort out rows where any crucial entry is not set
            if None in [row[i] for i in ROW_INDEXES.values()]:
                continue
            # Updating Groups and affiliated nodes from the first section
            cur_group = row[x1].strip()
            cur_node = (row[x2] + " " + row[x2 + 1]).strip().lower()
            if cur_group not in nx.get_node_attributes(self, "label").values():
                self.add_group(cur_group, i.g())
            cur_group_id = self.get_id_from_name(cur_group)
            if self.node_not_exists_in_group(cur_node, cur_group_id):
                cur_node_id = i.g()
                self.add_node(cur_node_id, label=cur_node, gid=cur_group_id, 
                               id=cur_node_id, graphics={"type": "ellipse"})
    
            # Updating Groups and affiliated nodes from the cause section
            cause_group = row[x3].strip()
            cause_node = row[x4].strip().lower()
            if cause_group not in nx.get_node_attributes(self, "label").values():
                self.add_group(cause_group, i.g())
            cause_group_id = self.get_id_from_name(cause_group)
            if self.node_not_exists_in_group(cause_node, cause_group_id):
                cause_node_id = i.g()
                self.add_node(cause_node_id, label=cause_node, gid=cause_group_id,
                               id=cause_node_id, graphics={"type": "ellipse"})
    
            # Updating Groups and affiliated nodes from the implication
            implic_group = row[x5].strip()
            implic_node = row[x6].strip().lower()
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
                
            #Adding risk to implicit nodes
            node_risk = RISK_MAPPING[row[x7]]["id"]
            self.update_node_risk(implic_node_id, node_risk)
        
    def add_group(self, name, group_id):
        """
        This adds a groups which just means to add a node with the 
        isGroup attribute set to 1
        """
        self.add_node(group_id, label=name, id=group_id, isGroup=1)

    def node_not_exists_in_group(self, node, group_id):
        """ Checks if the node is in the group with the given group_id """
        for node_id in self.nodes():
            if self.node[node_id]["label"] == node:
                if self.node[node_id]["gid"] == group_id:
                    return False
        return True

    def get_id_from_name(self, name, group=None):
        """
        This returns the id of a node or a group by giving the related id.
        If you want to check a node, the name of the affiliated group is 
        also needed.
        """
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
        """
        This sets the risk attribute for a node to the given risk.
        If the node already has a higher risk then it remains unchanged.
        """
        if ((not self.node[node_id].get("risk") or 
             self.node[node_id]["risk"] < node_risk)):
            self.node[node_id]["risk"] = node_risk

    def set_backdated_risks(self):
        """
        This updates the risk for all predecessors of the given node.
        """
        for node_id in nx.get_node_attributes(self, "risk"):
            node_risk = self.node[node_id]["risk"]
            predecessors = self.get_all_predecessors(node_id)
            if predecessors:
                for predecessor in predecessors:
                    self.update_node_risk(predecessor, node_risk)
            
    def get_all_predecessors(self, node_id):
        """
        Returns a set with all predecessors of the given node.
        """
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
        """
        Returns a set with all successors of the given node.
        """
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
        """
        Parse args strings given with -s and find all predecessors and
        successors and remove all other nodes from graph.
        """
        node_id = self.get_id_from_name(node_name, group_name)
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
        """
        Set the node colour depending on the nodes risk to the colour 
        which is defined in RISK_MAPPING.
        """
        for node_id in self.nodes():
            node_risk = self.node[node_id].get("risk")
            if node_risk:
                self.node[node_id]["graphics"]["fill"] = risk_colour(node_risk)

    def colour_edges(self):
        """
        Set the edge colour depending on the nodes risk to the colour 
        which is defined in RISK_MAPPING. Also the width of the edge is
        changed: The higher the risk, the thicker the edge.
        """
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
        """
        This removes all nodes from the graph which have lower risk than
        the desired risk.
        """
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
        """
        This clears edges which remain after some node where removed.
        They would not appear in the graph anyways but they appear in the 
        .gml file
        """
        nodes = self.nodes()
        for edge in self.edges():
            source = edge[0]
            target = edge[1]
            if source not in nodes or target not in nodes:
                self.remove_edge(source, target)


def get_xls_rows(file_, tab):
    """
    Loads the hazop table and reads it starting from the row defined by
    STARTING_ROW.
    This returns a list with specific openpyxl objects which you can access
    like a list with [].
    """
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
    """ Returns the affiliated colour for the given risk_id (e.g. "A")"""
    for risk in RISK_MAPPING.values():
        if risk_id in risk.values():
            return risk["colour"]

def main(args):
    rows = get_xls_rows(args.file, args.tab)
    
    graph = HazopGraph(rows)
    graph.create_graph_from_rows()
    if args.single_node:
        node = args.single_node[0].replace("_", " ")
        group = args.single_node[1].replace("_", " ")
        graph.get_single_node(node, group)
    graph.set_backdated_risks()
    graph.limit_risk_range(args.risk) # Default is set to "F"
    graph.colour_nodes()
    graph.colour_edges()
    graph.remove_remaining_edges()
    
    # Use modified version of write_gml to write the gml.
    gmlc.write_gml(graph, args.output)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-o', '--output', type=str, default="gmlcreator_test.gml",
                        help='path to target gmlfile')
    parser.add_argument('-r', '--risk', default="F",
                        help='minimum risk level to show (F to A)')
    parser.add_argument('-s', '--single-node', nargs="+", type=str,
                        help='View a single node and all its successors \
                        and predecessors. Syntax: -s no_flow C2')
    parser.add_argument('-f', '--file', type=str, default='HAZOP-Beispiel.xlsm',
                        help='input hazop xls/xlsm file')
    parser.add_argument('-t', '--tab', default='mHAZOP - Module',
                        help='name of the tab to use')

    args = parser.parse_args()
    main(args)
