source("FIPEX.output.to.R.input.r")
FiPEX.output.to.R.input()
source("dci.fxs.r")
x = try(dci.fxs(all.sections=T),silent=TRUE)
if(class(x)=='data.frame' | class(x)=='list'){
  write.table(x,file='out.txt')
} else{
  write("ERROR",file='out.txt')
}


