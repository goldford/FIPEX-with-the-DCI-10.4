dci.calc.all.sections.fx<-function(sum.table,lengths){

	#sum.table is a variable that is used in the dci.calc.r function, where we describe sum.table as sum.table.all or sum.table.nat
	#where: sum.table.nat<-read.csv("summary table natural.csv") and sum.table.all<-read.csv("summary table all.csv") - from the graph.and.data.setup.for.DCI.r function

	
#### DCI Calculation ####

#WHAT THIS FUNCTION DOES: it calculates the DCIp and DCId values for the riverscape (includes both natural and artificial barriers)






# calculate a DCI value using each section as the origin.	
	
#DCId calculation

sections<-as.vector(unique(sum.table$start))
# store the all section results in DCI.as
DCI.as<-NULL
for(s in 1:length(sections))
	{
	#browser()	
	DCId<-0	
	# select out only the data that corresponds to pathways from one sectino to all other sections
	d.nrows<-subset(sum.table, start==sections[s])
	d.sum.table<-d.nrows
	for (a in 1:dim(d.nrows)[1])
		{
		#to get the DCI for diadromous fish, use the following formula: 
		# DCId= li/L*Cj (where j= the product of the passability in the pathway)
	
		la<-d.sum.table$finish.section.length[a]/sum(lengths$Shape_Length)
		pass.d<-d.sum.table$pathway.pass[a]
		DCI.as[s]<-DCId+la*pass.d*100
		} # end loop over sections for dci calc
	print(DCI.as)	
	
	} # end loop over "first" sections
browser()
res<-data.frame(sections,DCI.as)


return(round(res,3))

}