dci.calc.fx<-function(sum.table,lengths,all.sections=F){

	#sum.table is a variable that is used in the dci.calc.r function, where we describe sum.table as sum.table.all or sum.table.nat
	#where: sum.table.nat<-read.csv("summary table natural.csv") and sum.table.all<-read.csv("summary table all.csv") - from the graph.and.data.setup.for.DCI.r function
	
	
#### DCI Calculation ####

#WHAT THIS FUNCTION DOES: it calculates the DCIp and DCId values for the riverscape (includes both natural and artificial barriers)


#this contains the length of each section

p.nrows<-dim(sum.table)[1]
#use this number for the DCIp calculation - interested in movements in all directions from all segments


d.nrows<-subset(sum.table, start=="sink")
#for diadromous fish we are only interested in the movement from the segment which is closest to the ocean
d.sum.table<-d.nrows

DCIp<-0
DCId<-0

#DCIp calculation
for (k in 1:p.nrows)
	{
		#to get the riverscape connectivity index for potadromous fish, use the given formula: DCIp= Cij*(li/L)*(lj/L)
		#Cij = passability for pathway (product of all barrier passabilities in the pathway), li & lj = length of start and finish sections, L = total length of all sections
	
		lj<-sum.table$start.section.length[k]/sum(lengths$Shape_Length)
		lk<-sum.table$finish.section.length[k]/sum(lengths$Shape_Length)
		pass<-sum.table$pathway.pass[k]
		DCIp<-DCIp+lj*lk*pass*100
	
		#add DCIp at the beginning to keep a running total of DCIp values
	}

#DCId calculation
for (a in 1:dim(d.nrows)[1])
	{
		#to get the DCI for diadromous fish, use the following formula: 
		# DCId= li/L*Cj (where j= the product of the passability in the pathway)
		
		la<-d.sum.table$finish.section.length[a]/sum(lengths$Shape_Length)
		pass.d<-d.sum.table$pathway.pass[a]
		DCId<-DCId+la*pass.d*100
	}

DCI<-t(c(DCIp,DCId))
DCI<-as.data.frame(DCI)	

names(DCI)<-c("DCIp","DCId")



#########  ALL SECTION ANLAYSIS  ######
## if desired, one can calculate the DCI_d starting with every sections.  This
## gives a "section-level" DCI score for each section in the watershed
if(all.sections==T)
	{

	sections<-as.vector(unique(sum.table$start))
	# store the all section results in DCI.as
	DCI.as<-NULL
	for(s in 1:length(sections))
		{

		DCI.s<-0	
		# select out only the data that corresponds to pathways from one sectino to all other sections
		d.nrows<-subset(sum.table, start==sections[s])
		d.sum.table<-d.nrows
		for (a in 1:dim(d.nrows)[1])
			{
			#to get the DCI for diadromous fish, use the following formula: 
			# DCId= li/L*Cj (where j= the product of the passability in the pathway)
		
			la<-d.sum.table$finish.section.length[a]/sum(lengths$Shape_Length)
			pass.d<-d.sum.table$pathway.pass[a]
			DCI.s<-DCI.s+la*pass.d*100
			} # end loop over sections for dci calc
			DCI.as[s]<-DCI.s	
		} # end loop over "first" sections	
		
	# STORE RESULTS IN .CSV file
	res<-data.frame(sections,DCI.as)	
	write.table(x=res,file="DCI.all.sections.csv",sep=",",row.names=F)
	
	} # end if statement over all.sections

#print(DCI)

#write.table(DCI,"DCI.csv", row.names=F, sep=",")

return(DCI)
#returns the results (but you can't do anything after this, so "return" must always be at the end of a function)


}