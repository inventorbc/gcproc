library(GCalignR)

# read in arguments after trailing arguments as vector, print to console, and check if at least one argument is supplied (input file).
args <- commandArgs(trailingOnly = TRUE)
cat(args, sep="\n")

if (length(args) == 0) {
	stop("Usage: RScript gcproc.R <path to input file>")
}

# Specify data location from first argument
peak_data <- args[1]

# Align the chromatograms from the input file. The reference is chosen as the first sample column "peaks", which is the list of analytes of interest.
peak_data_aligned <- align_chromatograms(data = peak_data,
				    rt_col_name = "RT",
				    reference = "peaks",
				    max_linear_shift = 0.05,
				    max_diff_peak2mean = 0.03,
				    min_diff_peak2peak = 0.03)

# Aligned data are stored in peak_data_aligned
print(peak_data_aligned$aligned$RT)
print(peak_data_aligned$aligned$Area)

# Find the sample properties - names and number of samples
sample_names <- colnames(peak_data_aligned$aligned$RT)[- (1:2)]
print(sample_names)
sample_num <- length(sample_names)

# Determine the peak indices for the analytes of interest 
peak_groups <- peak_data_aligned$aligned$RT[,2]
analyte_peaks <- (peak_groups[peak_groups != 0])
analyte_index <- which(peak_data_aligned$aligned$RT[,2] > 0, arr.ind=TRUE)

# Initialize analyte area matrix and populate with areas from each sample in the order of the peak list given. Convert matrix to dataframe.
analyte_areas <- matrix(nrow=sample_num, ncol=length(analyte_index))

for (i in 1:sample_num) {
	sample_peaks <- peak_data_aligned$aligned$RT[, i+2]
	sample_analyte_peaks <- sample_peaks[analyte_index]
		
	sample_areas <- peak_data_aligned$aligned$Area[, i+2]
	sample_analyte_areas <- sample_areas[analyte_index]
	
	analyte_areas[i,] <- sample_analyte_areas
}

analyte_areas <- data.frame(analyte_areas)
rownames(analyte_areas) <- sample_names

print("START Analyte Areas:")
print(analyte_areas)
print("END Analyte Areas:")

