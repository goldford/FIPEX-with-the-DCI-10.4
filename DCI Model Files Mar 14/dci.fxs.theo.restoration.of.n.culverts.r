###This function will give you the DCI values of the system where x number of culverts have been restored (i.e. passability of unnatural barriers coerced to 1)###

#n="__" are the number of culverts you want to restore
dci.fxs.theo.restoration.of.n.culverts<-function(n, passability)
{
	#source in functions into R
	#see each function for an explanation of inputs and outputs
	
	library(Rgraphviz)
	library(RBGL)
	library(combinat)
	
	source("convert.gis.output.to.r.format.r")
	source("get.adj.matrix.from.gis.r")
	source("graph.and.data.setup.for.DCI.r")
	source("sum.fx.r")
	source("dci.calc.fx.r")
	source("dci.calc.r")
	
	convert.gis.output.to.r.format()
	passability<-read.csv("segments.and.barriers.csv")
	#this has to be read in because "passability" doesn't get named until the "graph.and.data.setup.for.DCI.r" function gets called upon
	
	NB<-sum(passability$nat_barrier==TRUE)>0
	
	#get adjacency matrix information
	adj.matrix<-get.adj.matrix.from.gis()
	
	#we don't want to write over the original passability data in segments.and.barriers.csv - i.e. we don't overwrite the old values permanently
	passability.old<-passability
	
	dci.theo.n.culverts<-NULL
	barriers.restored<-NULL
	
	
	# GETTING ALL OF THE POSSIBLE UNIQUE PERMUTATIONS

	
	#want to find all of the unique FALSE barriers (i.e. all artificial barriers)
	barriers.artificial<-subset(passability,nat_barrier==FALSE)
	barriers.artificial.unique<-as.vector(unique(barriers.artificial$Bar_ID))
		
	
	#want to find all of the unique combinations of pairs, triplets, quads,... 
	all.possible.combinations<-combn(barriers.artificial.unique,n,fun=NULL, simplify=TRUE)
	#the output is a matrix
	#transpose the matrix, so that barriers combinations (each barrier is it's own cell), are read horizontally as opposed to vertically

	
	all.possible.combinations<-t(all.possible.combinations)
	
	#make it into a dataframe (it is still a matrix)
	all.possible.combinations<-as.data.frame(all.possible.combinations)
	
	#create a new column that pastes the row together  - NOT NEEDED BUT GOOD CODE TO KNOW
	#all.possible.combinations$combination<-do.call(paste,c(data.frame(all.possible.combinations),list(sep=',')))
	
	#need to now change the passabilities of the barrier combinations to "1" in the passability object 
	for(i in 1:length(all.possible.combinations[,1]))
	{
		#print("test")
		cat(i," of ", length(all.possible.combinations[,1]), " iterations \n")
		
		passability<-passability.old
		
		#need to match the barriers in each of the all.possible.combinations rows
		#all.possible.combiations seems to get messed up somehow - perhaps the using the "combn" function - so you have to double transpose it in order for R to read it in correctly		

		match.barriers<-match(passability$Bar_ID,t(t(all.possible.combinations[i,])))
		
		#assign the restored barriers a passability of 1
		passability$Pass[!is.na(match.barriers)]<-1
	
		#get the ID of the barriers that were restored
		barriers.restored.all<-as.vector(all.possible.combinations[i,])

  		#create a summary table for which we need to calculate the DCIs
	  	#the table will output start and end segments, all pathways, barriers encountered, passability, and start and finish segment lengths
            	        	
  		graph.and.data.setup.for.DCI(adj.matrix=adj.matrix,passability=passability,lengths=lengths)
            	
 	 	#calculate the DCI values and only store these values for the all barriers (i.e. DCIap and DCIad) - we aren't interested in the other outputs (i.e. DCInp and DCInd)
  		results<-dci.calc(NB)[1:2]
  	
  		results<-as.data.frame(results)
            	
  		#because we >return(DCI) in the dci.calc(), we can keep a running list of the results
  		dci.theo.n.culverts<-rbind(dci.theo.n.culverts,results)
  		barriers.restored<-rbind(barriers.restored, barriers.restored.all)
	}
 	#browser()
 	row.names(dci.theo.n.culverts)<-NULL
  	dci.theo.n.culverts<-data.frame(barriers.restored,dci.theo.n.culverts,row.names=NULL)
  	write.table(dci.theo.n.culverts,"dci.theo.n.culverts.csv", row.names=F,sep=",")

	
}
