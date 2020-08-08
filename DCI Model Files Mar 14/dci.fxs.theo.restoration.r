### This function allows you to change the value of the passability of 1 barrier at a time to see the change in the DCIp and DCId values.  In this function the passability is changed to "1" (see line 44 if you want to change the value).###

dci.fxs.theo.restoration<-function()
{
	#source in functions into R
	#see each function for an explanation of inputs and outputs
	
	##library(Rgraphviz)
	library(RBGL)
	
	source("convert.gis.output.to.r.format.r")
	source("get.adj.matrix.from.gis.r")
	source("graph.and.data.setup.for.DCI.r")
	source("sum.fx.r")
	source("dci.calc.fx.r")
	source("dci.calc.r")

	convert.gis.output.to.r.format()
	passability<-read.csv("segments.and.barriers.csv")
	#have to read it in because "passability" doesn't get named until the "graph.and.data.setup.for.DCI.r" function gets called upon
	
	#get adjacency information (only needs to be done once)	
	adj.matrix<-get.adj.matrix.from.gis()
	passability.old<-passability
	#we want to make sure that we preserve the original barrier data (don't overwrite old values permanently)
	NB<-sum(passability$nat_barrier==TRUE)>0
	
	dci.theo<-NULL	
	
	#cycle over unique barrier ID
	Bar_ID.unique<-unique(passability$Bar_ID)
	
	#for (i in 1:length(passability$Pass))
	for (i in 1:length(Bar_ID.unique))
	{

		#select the data for the unique barrier [i]. Select the first row, since each barrier appears twice
		unique.bar<-passability[passability$Bar_ID==Bar_ID.unique[i],][1,]
		
		#want the function to only look at the data with artificial barriers - coerce their passability to '1'
		if(unique.bar$nat_barrier==FALSE)	
		{
			passability<-passability.old
			passability$Pass[passability$Bar_ID==Bar_ID.unique[i]]<-1
			write.table(passability, "segments.and.barriers.csv", row.names=F, sep=",")
			#create the file "segment.and.barriers.csv" so that it gets called upon in "graph.and.data.setup.for.DCI.r" function
		
			Bar_ID<-subset(passability, nat_barrier==F)
			Bar_ID<-as.vector(Bar_ID$Bar_ID[Bar_ID$nat_barrier==F])
			Bar_ID<-as.vector(Bar_ID[1:(length(Bar_ID)/2)])
			
			# create summary table which we need to calculate the DCIs
			graph.and.data.setup.for.DCI(adj.matrix=adj.matrix,passability=passability)
	
			# Calculate the DCI and only store the values for the all barriers (the first two)
			results<-dci.calc(NB)[1:2]
			#because we >return(DCI) in the dci.calc(), we can keep a running list of the results  
			dci.theo<-rbind(dci.theo, results)
			#if(i==6) browser()
			#print(i)
			#print(passability$nat_barrier[i])
			#print(passability$Bar_ID[i])	
		}
	}

row.names(dci.theo)<-NULL
dci.theo<-cbind(Bar_ID,dci.theo)

write.table(dci.theo,"dci.theo.csv", row.names=F, sep=",")
#print(dci.theo)

}