library(xlsx)
#Jurkat Parse


fn = "/Users/Matthew/Documents/Work/Excel_Files/Jurkat_Qualitymetric_Report.xlsx.xlsx"



"Fucking stop this madness"


name = fn
operator= str(\)

parsedname = str_split(string = fn, pattern = "" )







wb = createWorkbook(type = "xlsx")
sheet1 = createSheet(wb=wb , sheetName ="Full_Report")
sheet2 = createSheet(wb=wb , sheetName ="Instrument_Report")


d = read.delim("V:/JurkatQC/Jurkat_Tesla/20151202/qualityMetricsExport.1.ssv", sep = ";")

d2= d[1,]

d3 = d2[,c("File","MS.MS.spectra.valid.V","Distinct.Peps.CS.Total....","Time.span.mid.90..matched.spectra.in.run..min.",
           "Gradient.Shape.mid.90..filtered.spectra.in.run","Median.MS1.peak.width.mid.90..matched.spectra..sec.",
           "Median.MS2.fill.time.mid.90..matched.spectra..msec.")]

row.names(d3) = NULL

addDataFrame(x = d, sheet=sheet1, row.names = FALSE)
addDataFrame(x = d3, sheet = sheet2, row.names = FALSE)
saveWorkbook(wb, "V:/JurkatQC/Jurkat_Tesla/20151202/qualityMetricsExpor.xlsx")