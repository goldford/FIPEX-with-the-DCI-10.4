test<-function(data=read.csv("FiPEX_connectivity.csv")){


	
####################################################################################
#####################################################################################
# create the file segment_matrix.csv that has columns: "Seg_ID","Seg"
# this file 
# read in a table of barriers with the downstream barrier listed.
# file name is "FiPEX_connectivity.csv", and is stored as the dataframe caleld "data"
# recreate a table with all the neighbours of each barrier listed. 


# the barrier names are stored in the BarrierOrFlagID column

# the closest downstream barrier is stored in Downstream_Barrier 

barrier<- as.vector(data$BarrierOrFlagID)
down.barrier<- as.vector(data$Downstream_Barrier)

# turn the barrier names into segment names by adding _s
segment<-paste(as.vector(data$BarrierOrFlagID),"_s",sep="")
down.segment<-paste(as.vector(data$Downstream_Barrier),"_s",sep="")

# maintian the name of the closest segment to the start point of the riverscape as "sink"
down.segment[down.segment=="Sink_s"]<-"sink"

res<-NULL

for(i in 1:length(segment))
	{
	# get the segment downstream	
	down<-down.segment[segment==segment[i]]
	
	# now look for any matching segments in the downstream dataset
	additional.connected.segments<-segment[!is.na(match(down.segment,segment[i]))]
	connected.segments<-c(segment[i],down,additional.connected.segments)
	
	num.segments<-length(connected.segments)
	newdata<-data.frame(rep(segment[i],num.segments),connected.segments)
	# make the names match what the R DCI functions expect
	names(newdata)<-c("Seg_ID","Seg")
	res<-rbind(res,newdata)
		
	}

# add in the origin, which is connected to itself and other segments that abut it. 

section.name<-"sink"
additional.connected.segments<-segment[!is.na(match(down.segment,section.name))]
connected.segments<-c(section.name,additional.connected.segments)
num.segments<-length(connected.segments)
newdata<-data.frame(rep(section.name,num.segments),connected.segments)
names(newdata)<-c("Seg_ID","Seg")
res<-rbind(res,newdata)

write.table(x=res,file="segment_matrix.csv",row.names=F,sep=",")

print(res)


#####################################################################################
#####################################################################################
# now create the file barrier.csv that has columns: Pass	Bar_ID	Seg_ID	nat_barrier
# this requires knowing the upstream and downstream segments of each barrier. 

# read in the barrier passabilities
barrier.info<-read.csv("FIPEX_BarrierHabitatLine.csv")
#barrier.info$BarrierID[barrier.info$BarrierID=="Sink"]<-"1"

# change the names to match the previous version
barrier.info$barrier<-barrier.info$BarrierID
barrier.info$pass<-barrier.info$BarrierPerm

res2<-NULL

for(i in 1:length(barrier))
	{
	upstream.segment<-paste(barrier[i],"_s",sep="")
	if(down.barrier[i]=="Sink") downstream.segment <- "sink"
	else downstream.segment<-paste(down.barrier[i],"_s",sep="")
	pass<-barrier.info$BarrierPerm[barrier.info$BarrierID==barrier[i]]
	nat_barrier<-barrier.info$NaturalYN[barrier.info$BarrierID==barrier[i]]
	#nat_barrier<-F
	
	newdata<-data.frame(rep(pass,2),rep(barrier[i],2),c(upstream.segment,downstream.segment),rep(nat_barrier,2))
	names(newdata)<-c("Pass","Bar_ID","Seg_ID","nat_barrier")
	res2<-rbind(res2,newdata)	
		
	}	
	
	
write.table(file="barrier.csv",res2,row.names=F,sep=",")	

print(res2)
####################################################################	
# now create the third file: lengths using the right segment names #
#######################################################################

data<-read.csv("FIPEX_BarrierHabitatLine.csv")
# the column names in the input file are: ObID	BarrierID	HabClass	Shape_Length	BarrierPerm	NaturalYN
# the column names in the destination file are: Seg_ID	Shape_Length

lengths<-data$Shape_Length
segment<-paste(as.vector(data$BarrierID),"_s",sep="")
# maintain the name of the closest segment to the start point of the riverscape as "sink"
segment[segment=="Sink_s"]<-"sink"


newdat<-data.frame(segment,lengths)
names(newdat)<-c("Seg_ID","Shape_Length")
write.table(file="length.csv",newdat,row.names=F,sep=",")

print(newdat)


}