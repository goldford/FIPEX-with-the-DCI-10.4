graph.fx<-function(edge.size=75,node.size=5,adj.matrix=read.csv("adjacency matrix.csv"),plot.it=F){

	#Problem with adjusting node.size.  If edge.size is changed, the font size on the graph changes, if the node.size is changed, nothing is changed on the graph.  
	#Note: the larger the edge.size number, the smaller the font.
	
	
#### INPUT DATA ####

#input: adj.matrix (created in "get.adj.matrix.from.gis.r")
#input: segments.and.barriers (created in "convert.gis.output.to.r.format.r") - segments.and.barriers is later called "passability"
#input: lengths (Excel file from GIS output - "length.csv")


##library(Rgraphviz)
library(RBGL)


#adjacency matrix.csv was created from the segment_matrix.csv file in the "get.adj.matrix.from.gis.r" function 
rownames(adj.matrix)<-1:length(adj.matrix)
colnames(adj.matrix)<-1:length(adj.matrix)

#need to convert the dataframe into a matrix
adj.matrix<-as.matrix(adj.matrix)

passability<-read.csv("segments.and.barriers.csv")
#this is the output from the "convert.gis.output.to.r.format.r" Crimson Editor file
#data contained in this file: "Bar_ID","Seg_1","Seg_2","Pass","nat_barrier", "section1.2" (pasting Seg_1 and Seg_2 together)

lengths<-read.csv("length.csv")
#contains the total length of each segment (i.e. segment id and segment length)
sections<-rownames(adj.matrix)

#to create a graph without arrows going from 1 to 1, or 2 to 2, or etc... you need to put zeros along the diagonal
adj.matrix.zeros.on.diag<-adj.matrix
diag(adj.matrix.zeros.on.diag)<-0


if(plot.it==T)
	{
#### CREATE GRAPH ####

g1<-new("graphAM",adjMat=adj.matrix.zeros.on.diag, edgemode="directed")
#for the graph, note that the adj.matrix.zeros.on.diag must have 0s across the diagonal or else it will give you more edges, because it would include a-a, b-b, c-c, etc...

#we want to label edges of the graph with the Barrier letter and the passability value
#need to create a new column in passability that gives you: Bar_ID (Passability)
#round the passability values to 2 decimal places (so it fits better on the page)
pass.barrier<-passability
pass.barrier$bar.pass<-with(pass.barrier,paste(Bar_ID," (",round(Pass,2),")",sep=""))

#deal with EDGES
eAttrs<-list()

#need to assign correct labeling in Graph
#create new labels in pass.barrier that match graph object labels (edgeNames(g1)/names(eAttrs$labels))
pass.barrier$names.eAttrs<-with(pass.barrier, paste(Seg_1, Seg_2, sep="~"))

#determine the right order:
#the "match" must be done in this order so that we know the position of where the 1st element of edgeNames(g1) matches with pass.barrier$names.eAttrs
ord<-match(edgeNames(g1),pass.barrier$names.eAttrs)
# order the dataframe appropriately wrt the ord and pass.barrier dataframes
pass.barrier<-data.frame(pass.barrier)[ord,]

#assign the text we want to appear with each edge
ew<-pass.barrier$bar.pass
# get the labels of the edges
names(ew)<-edgeNames(g1)
# assign graph object labels so it it knows where to put the text
eAttrs$label<-ew

#deal with NODES
nAttrs<-list()
#get segment names
n<-nAttrs$label
n<-row.names(adj.matrix)
names(n)<-nodes(g1)

# fontsize for edges 
a<-rep(edge.size,length(ew))
names(a)<-edgeNames(g1)
eAttrs$fontsize<-a

# fontsize for nodes 
b<-rep(node.size, length(nodes(g1)))
names(b)<-nodes(g1)
nAttrs$fontsize<-b

#overall it looks like font size of the nodes and edges don't work independently from eachother

#node shape
node.shape<-rep("ellipse",length(nodes(g1)))
names(node.shape)<-nodes(g1)
nAttrs$shape<-node.shape
#node height
node.height<-rep(4,length(nodes(g1)))
names(node.height)<-nodes(g1)
nAttrs$height<-node.height
#node width
node.width<-rep(1.1,length(nodes(g1)))
names(node.width)<-nodes(g1)
nAttrs$width<-node.width
#edge color
edge.color<-rep("grey",length(ew))
names(edge.color)<-edgeNames(g1)
eAttrs$color<-edge.color

plot(g1,edgeAttrs=eAttrs, nodeAttrs=nAttrs, main="_______ Watershed")
#gives a graph with a 2-way arrow
}

g2<-new("graphAM",adjMat=adj.matrix, edgemode="directed")
#use g2 for the graph.and.data.setup.for DCI function - you need 1's along the diagonal

#return(c(sections,g2))
}