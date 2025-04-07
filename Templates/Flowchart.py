from graphviz import Digraph

# Create a new Digraph
flowchart = Digraph("ICSS Flowchart", format="png")

# Define Nodes (Processes)
flowchart.node("Start", shape="ellipse", style="filled", fillcolor="lightgreen")
flowchart.node("Fetch Data", shape="box", style="filled", fillcolor="lightblue")
flowchart.node("Preprocess Data", shape="box", style="filled", fillcolor="lightblue")
flowchart.node("Classify Complaints", shape="diamond", style="filled", fillcolor="lightpink")
flowchart.node("Generate Alerts", shape="box", style="filled", fillcolor="yellow")
flowchart.node("Display Dashboard", shape="box", style="filled", fillcolor="orange")
flowchart.node("End", shape="ellipse", style="filled", fillcolor="red")

# Define Edges (Connections)
flowchart.edge("Start", "Fetch Data")
flowchart.edge("Fetch Data", "Preprocess Data")
flowchart.edge("Preprocess Data", "Classify Complaints")
flowchart.edge("Classify Complaints", "Generate Alerts", label="High Priority")
flowchart.edge("Classify Complaints", "Display Dashboard", label="Low/Moderate Priority")
flowchart.edge("Generate Alerts", "Display Dashboard")
flowchart.edge("Display Dashboard", "End")

# Render and save flowchart
flowchart.render("ICSS_Flowchart", view=True)  # Opens the flowchart automatically
