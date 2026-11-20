namespace MicroadsOffice.Excel.Formula;

public static partial class FunctionCatalog
{
    internal const string DataMisc = @"
ACOSH|M|number|Cos hyperbol ngược|Inverse hyperbolic cosine
ASINH|M|number|Sin hyperbol ngược|Inverse hyperbolic sine
ATANH|M|number|Tang hyperbol ngược|Inverse hyperbolic tangent
ACOT|M|number|Arccotang|Arccotangent
ACOTH|M|number|Cotang hyperbol ngược|Inverse hyperbolic cotangent
COT|M|number|Cotang của góc|Cotangent of an angle
COTH|M|number|Cotang hyperbol|Hyperbolic cotangent
CSC|M|number|Cosec của góc|Cosecant of an angle
CSCH|M|number|Cosec hyperbol|Hyperbolic cosecant
SEC|M|number|Sec của góc|Secant of an angle
SECH|M|number|Sec hyperbol|Hyperbolic secant
ARABIC|M|text|Đổi số La Mã sang số Ả Rập|Converts a Roman numeral to an Arabic numeral
BASE|M|number, radix, [min_length]|Đổi số sang hệ cơ số khác (dạng chữ)|Converts a number to text in the given radix
DECIMAL|M|text, radix|Đổi chuỗi số theo cơ số sang số thập phân|Converts text in a given base to a decimal number
COMBINA|M|number, number_chosen|Số tổ hợp có lặp|Number of combinations with repetitions
FACTDOUBLE|M|number|Giai thừa kép|Double factorial
MDETERM|M|array|Định thức ma trận|Matrix determinant
MINVERSE|M|array|Ma trận nghịch đảo|Inverse matrix
MMULT|M|array1, array2|Tích hai ma trận|Matrix product of two arrays
MUNIT|M|dimension|Ma trận đơn vị|Unit matrix of the given dimension
MULTINOMIAL|M|number1, [number2], ...|Hệ số đa thức|Multinomial of a set of numbers
RANDARRAY|M|[rows], [columns], [min], [max], [whole_number]|Mảng số ngẫu nhiên|Array of random numbers
SERIESSUM|M|x, n, m, coefficients|Tổng chuỗi lũy thừa|Sum of a power series
SUMX2MY2|M|array_x, array_y|Tổng hiệu các bình phương|Sum of the differences of squares
SUMX2PY2|M|array_x, array_y|Tổng tổng các bình phương|Sum of the sums of squares
SUMXMY2|M|array_x, array_y|Tổng bình phương các hiệu|Sum of squares of differences
PERCENTOF|M|data_subset, data_all|Tỷ lệ phần trăm của tập con trên tổng|Percentage a subset makes up of the total
AGGREGATE|M|function_num, options, ref1, ...|Tổng hợp (có thể bỏ qua hàng ẩn, giá trị lỗi)|Aggregate that can ignore hidden rows and errors
COVARIANCE.P|S|array1, array2|Hiệp phương sai tổng thể|Population covariance
COVARIANCE.S|S|array1, array2|Hiệp phương sai mẫu|Sample covariance
DEVSQ|S|number1, [number2], ...|Tổng bình phương độ lệch|Sum of squares of deviations
FREQUENCY|S|data_array, bins_array|Phân phối tần suất|Frequency distribution as a vertical array
GROWTH|S|known_y's, [known_x's], [new_x's], [const]|Dự báo tăng trưởng hàm mũ|Predicts exponential growth
KURT|S|number1, [number2], ...|Độ nhọn của phân phối|Kurtosis of a data set
LINEST|S|known_y's, [known_x's], [const], [stats]|Tham số hồi quy tuyến tính|Parameters of a linear trend
LOGEST|S|known_y's, [known_x's], [const], [stats]|Tham số hồi quy hàm mũ|Parameters of an exponential trend
MODE.MULT|S|number1, [number2], ...|Các giá trị xuất hiện nhiều nhất|Vertical array of the most frequent values
PERCENTRANK.EXC|S|array, x, [significance]|Thứ hạng phần trăm (loại trừ 0 và 1)|Percentage rank (0..1 exclusive)
PERCENTRANK.INC|S|array, x, [significance]|Thứ hạng phần trăm|Percentage rank (0..1 inclusive)
QUARTILE.EXC|S|array, quart|Tứ phân vị (loại trừ 0 và 1)|Quartile (0..1 exclusive)
RSQ|S|known_y's, known_x's|Bình phương hệ số tương quan|Square of the Pearson correlation coefficient
SKEW|S|number1, [number2], ...|Độ lệch của phân phối (mẫu)|Skewness of a distribution
SKEW.P|S|number1, [number2], ...|Độ lệch của phân phối (tổng thể)|Skewness of a population
STEYX|S|known_y's, known_x's|Sai số chuẩn của giá trị y dự báo|Standard error of the predicted y-values
TREND|S|known_y's, [known_x's], [new_x's], [const]|Dự báo theo xu hướng tuyến tính|Values along a linear trend
TRIMMEAN|S|array, percent|Trung bình sau khi bỏ các giá trị ở hai đầu|Mean of the interior of a data set
STDEVPA|S|value1, [value2], ...|Độ lệch chuẩn tổng thể (tính cả chữ, logic)|Population standard deviation including text and logicals
VARA|S|value1, [value2], ...|Phương sai mẫu (tính cả chữ, logic)|Sample variance including text and logicals
VARPA|S|value1, [value2], ...|Phương sai tổng thể (tính cả chữ, logic)|Population variance including text and logicals
COVAR|C|array1, array2|Hiệp phương sai (hàm cũ)|Covariance (legacy)
PERCENTRANK|C|array, x, [significance]|Thứ hạng phần trăm (hàm cũ)|Percentage rank (legacy)
ARRAYTOTEXT|T|array, [format]|Chuyển mảng thành chuỗi|Converts an array to text
VALUETOTEXT|T|value, [format]|Chuyển giá trị thành chuỗi|Converts a value to text
TEXTSPLIT|T|text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]|Tách chuỗi thành các cột, hàng|Splits text into columns and rows
REGEXTEST|T|text, pattern, [case_sensitivity]|Kiểm tra chuỗi khớp biểu thức chính quy|Checks whether text matches a regular expression
REGEXEXTRACT|T|text, pattern, [return_mode], [case_sensitivity]|Trích chuỗi khớp biểu thức chính quy|Extracts text matching a regular expression
REGEXREPLACE|T|text, pattern, replacement, [occurrence], [case_sensitivity]|Thay chuỗi khớp biểu thức chính quy|Replaces text matching a regular expression
ASC|T|text|Đổi ký tự toàn độ rộng sang nửa độ rộng|Changes full-width characters to half-width
DBCS|T|text|Đổi ký tự nửa độ rộng sang toàn độ rộng|Changes half-width characters to full-width
JIS|T|text|Đổi ký tự nửa độ rộng sang toàn độ rộng|Changes half-width characters to full-width
BAHTTEXT|T|number|Đọc số thành chữ tiếng Thái (baht)|Converts a number to Thai text with Baht
FINDB|T|find_text, within_text, [start_num]|Tìm (phân biệt hoa thường, theo byte)|Finds text (case-sensitive, byte based)
SEARCHB|T|find_text, within_text, [start_num]|Tìm (không phân biệt hoa thường, theo byte)|Finds text (not case-sensitive, byte based)
LEFTB|T|text, [num_bytes]|Ký tự bên trái (theo byte)|Leftmost characters (byte based)
RIGHTB|T|text, [num_bytes]|Ký tự bên phải (theo byte)|Rightmost characters (byte based)
MIDB|T|text, start_num, num_bytes|Ký tự ở giữa (theo byte)|Characters from the middle (byte based)
LENB|T|text|Độ dài chuỗi (theo byte)|Number of bytes in a text
REPLACEB|T|old_text, start_num, num_bytes, new_text|Thay ký tự theo vị trí (theo byte)|Replaces characters by position (byte based)
PHONETIC|T|reference|Trích phiên âm (furigana) của chuỗi|Extracts the phonetic (furigana) characters
NETWORKDAYS.INTL|D|start_date, end_date, [weekend], [holidays]|Số ngày làm việc (tùy chọn ngày nghỉ cuối tuần)|Working days with custom weekends
WORKDAY.INTL|D|start_date, days, [weekend], [holidays]|Ngày làm việc sau n ngày (tùy chọn ngày nghỉ cuối tuần)|Working day after n days with custom weekends
INFO|I|type_text|Thông tin về môi trường làm việc|Information about the operating environment
DAVERAGE|B|database, field, criteria|Trung bình các bản ghi thỏa điều kiện|Average of selected database entries
DCOUNT|B|database, [field], criteria|Đếm ô chứa số trong các bản ghi thỏa điều kiện|Counts cells containing numbers in a database
DCOUNTA|B|database, [field], criteria|Đếm ô không trống trong các bản ghi thỏa điều kiện|Counts non-blank cells in a database
DGET|B|database, field, criteria|Lấy một giá trị duy nhất thỏa điều kiện|Extracts a single record matching the criteria
DMAX|B|database, field, criteria|Giá trị lớn nhất của các bản ghi thỏa điều kiện|Maximum of selected database entries
DMIN|B|database, field, criteria|Giá trị nhỏ nhất của các bản ghi thỏa điều kiện|Minimum of selected database entries
DPRODUCT|B|database, field, criteria|Tích các giá trị thỏa điều kiện|Multiplies values of selected records
DSTDEV|B|database, field, criteria|Độ lệch chuẩn mẫu của các bản ghi thỏa điều kiện|Sample standard deviation of selected entries
DSTDEVP|B|database, field, criteria|Độ lệch chuẩn tổng thể của các bản ghi thỏa điều kiện|Population standard deviation of selected entries
DSUM|B|database, field, criteria|Tổng các giá trị thỏa điều kiện|Adds the values of selected records
DVAR|B|database, field, criteria|Phương sai mẫu của các bản ghi thỏa điều kiện|Sample variance of selected entries
DVARP|B|database, field, criteria|Phương sai tổng thể của các bản ghi thỏa điều kiện|Population variance of selected entries
ENCODEURL|W|text|Mã hóa chuỗi để dùng trong URL|URL-encodes a text
FILTERXML|W|xml, xpath|Lấy dữ liệu từ XML theo XPath|Returns data from XML using an XPath
";
}
