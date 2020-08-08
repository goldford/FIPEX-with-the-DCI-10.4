source("dci.fxs.r")
x = try(dci.fxs(),silent=TRUE)
if(class(x)=='data.frame'){
  write.table(x,file='out.txt')
} else{
  write("ERROR",file='out.txt')
}


