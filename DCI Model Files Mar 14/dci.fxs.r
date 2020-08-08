dci.fxs<-function(all.sections=F){

	
#source in the 7 functions in R
#see each function for an explanation of inputs and outputs

source("convert.gis.output.to.r.format.r")
source("get.adj.matrix.from.gis.r")
source("graph.fx.r")
source("sum.fx.r")
source("graph.and.data.setup.for.DCI.r")
source("dci.calc.fx.r")
source("dci.calc.r")


convert.gis.output.to.r.format()
adj.matrix<-get.adj.matrix.from.gis()
#have to assign it an object name as this function gets called in in "graph.and.data.setup.for.DCI.r" and "sum.fx.r"

graph.fx(plot.it=F)
passability<-read.csv("segments.and.barriers.csv")

NB<-graph.and.data.setup.for.DCI(passability=passability, adj.matrix=adj.matrix)
dci.calc(NB,all.sections=all.sections)



}