graph.and.data.setup.for.DCI<-function(adj.matrix,passability,lengths){

#dci.fxs.r assigns adj.matrix and passability before it runs the graph.and.data.setup.for.DCI.r	
#output: get the summary tables of all possible pathways with their passability values, start and end sections, barrier id, length of start and end segments
#calls on "sum.fx.r" function



#### FUNCTION ####

#output: this creates 2 tables.  "summary table natural.csv"  contains the passability values of each pathway in the riverscape if there were NO artificial barriers (the natural passability of the riverscape) and "summary table all.csv" contains the passability values of the pathways in the given riverscape (artificial and natural barriers included).  This way we can have an idea of how much the artificial barriers are really affecting the DCI of the riverscape.
#output: start and end segments, pathway, barrier.id, length of start and finish sections

NB<-sum(passability$nat_barrier==TRUE)>0
#NB = natural barriers
if (NB==TRUE)
	{
	#if there are natural barriers in the system, we want to know what the overall DCInp and DCInd (n= natural) is for the system so we can compare it to the DCIap and DCIad (a= all (anthropogenic + natural barriers))
	natural.passability<-passability
	#because we want to know what the natural DCI value is, we only want to take into consideration passability values for natural barriers.  so we want to change the passability values for anthropogenic barriers to 1
	natural.passability$Pass[natural.passability$nat_barrier==FALSE]<-1

	sum.table.n<-sum.fx(adj.matrix=adj.matrix, passability=natural.passability,lengths=lengths)
	write.table(sum.table.n,"summary table natural.csv",row.names=F,sep=",")
	}

#if you put the else statement here, then it will only do the below commands only if the above statement is false.  if you take the else statement out, it will do the above commands only if it's true, but it will ALWAYS do the commands below.


sum.table<-sum.fx(adj.matrix=adj.matrix, passability=passability,lengths=lengths)

write.table(sum.table,"summary table all.csv",row.names=F,sep=",")

return(NB)
}