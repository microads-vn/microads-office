namespace MicroadsOffice.Excel.Formula;

public static partial class FunctionCatalog
{
    /// <summary>Probability distributions and hypothesis tests (S) and their Excel 2007 compatibility names (C).</summary>
    internal const string DataDist = @"
BETA.DIST|S|x, alpha, beta, cumulative, [A], [B]|Phân phối beta|Beta distribution
BETA.INV|S|probability, alpha, beta, [A], [B]|Hàm ngược của phân phối beta tích lũy|Inverse of the cumulative beta distribution
BINOM.DIST|S|number_s, trials, probability_s, cumulative|Xác suất phân phối nhị thức|Binomial distribution probability
BINOM.DIST.RANGE|S|trials, probability_s, number_s, [number_s2]|Xác suất số lần thành công nằm trong khoảng (nhị thức)|Probability of a trial result falling in a range (binomial)
BINOM.INV|S|trials, probability_s, alpha|Giá trị nhỏ nhất có xác suất nhị thức tích lũy ≥ alpha|Smallest value whose cumulative binomial probability is ≥ alpha
CHISQ.DIST|S|x, deg_freedom, cumulative|Phân phối khi bình phương|Chi-squared distribution
CHISQ.DIST.RT|S|x, deg_freedom|Xác suất đuôi phải của phân phối khi bình phương|Right-tailed probability of the chi-squared distribution
CHISQ.INV|S|probability, deg_freedom|Hàm ngược của phân phối khi bình phương (đuôi trái)|Inverse of the left-tailed chi-squared distribution
CHISQ.INV.RT|S|probability, deg_freedom|Hàm ngược của phân phối khi bình phương (đuôi phải)|Inverse of the right-tailed chi-squared distribution
CHISQ.TEST|S|actual_range, expected_range|Kiểm định khi bình phương về tính độc lập|Chi-squared test for independence
CONFIDENCE.NORM|S|alpha, standard_dev, size|Khoảng tin cậy của trung bình (phân phối chuẩn)|Confidence interval for a mean (normal distribution)
CONFIDENCE.T|S|alpha, standard_dev, size|Khoảng tin cậy của trung bình (phân phối Student)|Confidence interval for a mean (Student's t)
EXPON.DIST|S|x, lambda, cumulative|Phân phối mũ|Exponential distribution
F.DIST|S|x, deg_freedom1, deg_freedom2, cumulative|Phân phối F|F probability distribution
F.DIST.RT|S|x, deg_freedom1, deg_freedom2|Xác suất đuôi phải của phân phối F|Right-tailed F probability distribution
F.INV|S|probability, deg_freedom1, deg_freedom2|Hàm ngược của phân phối F|Inverse of the F probability distribution
F.INV.RT|S|probability, deg_freedom1, deg_freedom2|Hàm ngược của phân phối F đuôi phải|Inverse of the right-tailed F distribution
F.TEST|S|array1, array2|Kiểm định F (so sánh hai phương sai)|Result of an F-test (compares two variances)
FISHER|S|x|Phép biến đổi Fisher|Fisher transformation
FISHERINV|S|y|Nghịch đảo phép biến đổi Fisher|Inverse of the Fisher transformation
GAMMA|S|number|Hàm gamma|Gamma function value
GAMMA.DIST|S|x, alpha, beta, cumulative|Phân phối gamma|Gamma distribution
GAMMA.INV|S|probability, alpha, beta|Hàm ngược của phân phối gamma tích lũy|Inverse of the gamma cumulative distribution
GAMMALN|S|x|Lôgarit tự nhiên của hàm gamma|Natural logarithm of the gamma function
GAMMALN.PRECISE|S|x|Lôgarit tự nhiên của hàm gamma|Natural logarithm of the gamma function
GAUSS|S|z|Xác suất chuẩn tắc nằm giữa trung bình và z|Probability between the mean and z standard deviations
HYPGEOM.DIST|S|sample_s, number_sample, population_s, number_pop, cumulative|Phân phối siêu bội|Hypergeometric distribution
LOGNORM.DIST|S|x, mean, standard_dev, cumulative|Phân phối lôga chuẩn|Lognormal distribution
LOGNORM.INV|S|probability, mean, standard_dev|Hàm ngược của phân phối lôga chuẩn|Inverse of the lognormal distribution
NEGBINOM.DIST|S|number_f, number_s, probability_s, cumulative|Phân phối nhị thức âm|Negative binomial distribution
NORM.DIST|S|x, mean, standard_dev, cumulative|Phân phối chuẩn|Normal distribution
NORM.INV|S|probability, mean, standard_dev|Hàm ngược của phân phối chuẩn tích lũy|Inverse of the normal cumulative distribution
NORM.S.DIST|S|z, cumulative|Phân phối chuẩn tắc|Standard normal distribution
NORM.S.INV|S|probability|Hàm ngược của phân phối chuẩn tắc tích lũy|Inverse of the standard normal cumulative distribution
PERMUTATIONA|S|number, number_chosen|Số chỉnh hợp lặp|Number of permutations with repetitions
PHI|S|x|Mật độ của phân phối chuẩn tắc|Density of the standard normal distribution
POISSON.DIST|S|x, mean, cumulative|Phân phối Poisson|Poisson distribution
PROB|S|x_range, prob_range, lower_limit, [upper_limit]|Xác suất giá trị nằm trong khoảng|Probability that values are between two limits
STANDARDIZE|S|x, mean, standard_dev|Giá trị chuẩn hóa (z)|Normalized value (z-score)
T.DIST|S|x, deg_freedom, cumulative|Phân phối Student (t)|Student's t-distribution
T.DIST.2T|S|x, deg_freedom|Phân phối Student hai phía|Two-tailed Student's t-distribution
T.DIST.RT|S|x, deg_freedom|Phân phối Student đuôi phải|Right-tailed Student's t-distribution
T.INV|S|probability, deg_freedom|Hàm ngược của phân phối Student (đuôi trái)|Left-tailed inverse of Student's t-distribution
T.INV.2T|S|probability, deg_freedom|Hàm ngược của phân phối Student hai phía|Two-tailed inverse of Student's t-distribution
T.TEST|S|array1, array2, tails, type|Kiểm định t của Student|Probability associated with a Student's t-test
WEIBULL.DIST|S|x, alpha, beta, cumulative|Phân phối Weibull|Weibull distribution
Z.TEST|S|array, x, [sigma]|Giá trị P một phía của kiểm định z|One-tailed P-value of a z-test
BETADIST|C|x, alpha, beta, [A], [B]|Phân phối beta tích lũy|Cumulative beta distribution
BETAINV|C|probability, alpha, beta, [A], [B]|Hàm ngược của phân phối beta tích lũy|Inverse of the cumulative beta distribution
BINOMDIST|C|number_s, trials, probability_s, cumulative|Xác suất phân phối nhị thức|Binomial distribution probability
CHIDIST|C|x, deg_freedom|Xác suất đuôi phải của phân phối khi bình phương|Right-tailed probability of the chi-squared distribution
CHIINV|C|probability, deg_freedom|Hàm ngược của phân phối khi bình phương (đuôi phải)|Inverse of the right-tailed chi-squared distribution
CHITEST|C|actual_range, expected_range|Kiểm định khi bình phương về tính độc lập|Chi-squared test for independence
CONFIDENCE|C|alpha, standard_dev, size|Khoảng tin cậy của trung bình (phân phối chuẩn)|Confidence interval for a population mean
CRITBINOM|C|trials, probability_s, alpha|Giá trị nhỏ nhất có xác suất nhị thức tích lũy ≥ alpha|Smallest value whose cumulative binomial probability is ≥ alpha
EXPONDIST|C|x, lambda, cumulative|Phân phối mũ|Exponential distribution
FDIST|C|x, deg_freedom1, deg_freedom2|Xác suất đuôi phải của phân phối F|Right-tailed F probability distribution
FINV|C|probability, deg_freedom1, deg_freedom2|Hàm ngược của phân phối F đuôi phải|Inverse of the right-tailed F distribution
FTEST|C|array1, array2|Kiểm định F (so sánh hai phương sai)|Result of an F-test
GAMMADIST|C|x, alpha, beta, cumulative|Phân phối gamma|Gamma distribution
GAMMAINV|C|probability, alpha, beta|Hàm ngược của phân phối gamma tích lũy|Inverse of the gamma cumulative distribution
HYPGEOMDIST|C|sample_s, number_sample, population_s, number_pop|Xác suất phân phối siêu bội|Hypergeometric distribution probability
LOGINV|C|probability, mean, standard_dev|Hàm ngược của phân phối lôga chuẩn|Inverse of the lognormal cumulative distribution
LOGNORMDIST|C|x, mean, standard_dev|Phân phối lôga chuẩn tích lũy|Cumulative lognormal distribution
NEGBINOMDIST|C|number_f, number_s, probability_s|Xác suất phân phối nhị thức âm|Negative binomial distribution probability
NORMDIST|C|x, mean, standard_dev, cumulative|Phân phối chuẩn|Normal distribution
NORMINV|C|probability, mean, standard_dev|Hàm ngược của phân phối chuẩn tích lũy|Inverse of the normal cumulative distribution
NORMSDIST|C|z|Phân phối chuẩn tắc tích lũy|Standard normal cumulative distribution
NORMSINV|C|probability|Hàm ngược của phân phối chuẩn tắc tích lũy|Inverse of the standard normal cumulative distribution
POISSON|C|x, mean, cumulative|Phân phối Poisson|Poisson distribution
TDIST|C|x, deg_freedom, tails|Phân phối Student (t)|Student's t-distribution
TINV|C|probability, deg_freedom|Hàm ngược của phân phối Student hai phía|Two-tailed inverse of Student's t-distribution
TTEST|C|array1, array2, tails, type|Kiểm định t của Student|Probability associated with a Student's t-test
WEIBULL|C|x, alpha, beta, cumulative|Phân phối Weibull|Weibull distribution
ZTEST|C|array, x, [sigma]|Giá trị P một phía của kiểm định z|One-tailed P-value of a z-test
";
}
