dci.calc<-function(NB,lengths=read.csv("length.csv"),sum.table.all=read.csv("summary table all.csv"),all.sections=F)
{
	

#### DCI Calculation ####

#calls upon the "dci.calc.fx.r function"

if (NB==T){
sum.table.nat<-read.csv("summary table natural.csv")
#this dataframe was created in the "graph.and.data.setup.for.DCI.r" Crimson Editor file
#dataframe includes start and end segments, pathway, barriers in pathway, passability for the pathway, and the length of the start and end segments
DCI.n<-dci.calc.fx(sum.table=sum.table.nat,lengths=lengths)
write.table(DCI.n,"DCIn.csv", row.names=F, sep=",")
}


#the summary table all.csv dataframe was created in the "graph.and.data.setup.for.DCI.r" Crimson Editor file
#it includes start and end segments, pathway, barriers in pathway, passability for the pathway, and the length of the start and end segments

DCI.a<-dci.calc.fx(sum.table=sum.table.all,lengths=lengths,all.sections=all.sections)
write.table(DCI.a,"DCIa.csv", row.names=F, sep=",")

if (NB==T){
#returns the results (but you can't do anything after this, so "return" must always be at the end of a function)
prop.of.DCI.n<-round(DCI.a/DCI.n,3)
write.table(prop.of.DCI.n,"prop.of.DCI.n.csv",row.names=F, sep=",")

res<-data.frame(unlist(c(DCI.a,DCI.n,prop.of.DCI.n)))
row.names(res)<-c("DCI_P (Total)","DCI_D (Total)","DCI_P (nat. barriers only)","DCI_D (nat. barriers only)","DCI_P (prop.of natural)","DCI_D (prop.of natural)")
names(res)<-"value"
return(res)
	}

else {
	res<- data.frame(t(DCI.a))
	names(res)<-"value"
	return(res)
	}


}