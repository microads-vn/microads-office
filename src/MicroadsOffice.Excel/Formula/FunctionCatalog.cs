using MicroadsOffice.Core.Office;

namespace MicroadsOffice.Excel.Formula;

public enum FnCategory { Math, Statistical, Logical, Lookup, Text, DateTime, Financial, Information, Engineering, Database, Compatibility, Web }

public sealed record FunctionInfo(string Name, FnCategory Category, string Args, string DescVi, string DescEn)
{
    public string Description => Loc.T(DescVi, DescEn);
    public string Signature => Name + "(" + Args + ")";
}

/// <summary>Names, categories, argument lists and descriptions (VI/EN) of every function.</summary>
public static partial class FunctionCatalog
{
    public static string CategoryName(FnCategory c) => c switch
    {
        FnCategory.Math => Loc.T("Toán học & lượng giác", "Math & Trig"),
        FnCategory.Statistical => Loc.T("Thống kê", "Statistical"),
        FnCategory.Logical => Loc.T("Logic", "Logical"),
        FnCategory.Lookup => Loc.T("Tra cứu & tham chiếu", "Lookup & Reference"),
        FnCategory.Text => Loc.T("Văn bản", "Text"),
        FnCategory.DateTime => Loc.T("Ngày & giờ", "Date & Time"),
        FnCategory.Financial => Loc.T("Tài chính", "Financial"),
        FnCategory.Engineering => Loc.T("Kỹ thuật", "Engineering"),
        FnCategory.Database => Loc.T("Cơ sở dữ liệu", "Database"),
        FnCategory.Compatibility => Loc.T("Tương thích (bản cũ)", "Compatibility"),
        FnCategory.Web => Loc.T("Web", "Web"),
        _ => Loc.T("Thông tin", "Information")
    };

    private const string Data = @"
SUM|M|number1, [number2], ...|Cộng tất cả các số|Adds all the numbers
SUMIF|M|range, criteria, [sum_range]|Cộng các ô thỏa một điều kiện|Adds the cells specified by a condition
SUMIFS|M|sum_range, criteria_range1, criteria1, ...|Cộng các ô thỏa nhiều điều kiện|Adds the cells that meet multiple criteria
SUMPRODUCT|M|array1, [array2], ...|Tổng các tích của các mảng tương ứng|Returns the sum of the products of corresponding array components
SUMSQ|M|number1, [number2], ...|Tổng bình phương|Sum of the squares
PRODUCT|M|number1, [number2], ...|Nhân các số|Multiplies its arguments
ROUND|M|number, num_digits|Làm tròn số|Rounds a number to a specified number of digits
ROUNDUP|M|number, num_digits|Làm tròn lên (xa số 0)|Rounds a number up, away from zero
ROUNDDOWN|M|number, num_digits|Làm tròn xuống (về số 0)|Rounds a number down, toward zero
MROUND|M|number, multiple|Làm tròn tới bội số gần nhất|Rounds to the desired multiple
TRUNC|M|number, [num_digits]|Cắt bỏ phần thập phân|Truncates a number
INT|M|number|Làm tròn xuống số nguyên|Rounds down to the nearest integer
MOD|M|number, divisor|Số dư của phép chia|Returns the remainder from division
QUOTIENT|M|numerator, denominator|Phần nguyên của phép chia|Integer portion of a division
ABS|M|number|Giá trị tuyệt đối|Absolute value
SQRT|M|number|Căn bậc hai|Positive square root
SQRTPI|M|number|Căn bậc hai của (số × pi)|Square root of (number * pi)
POWER|M|number, power|Lũy thừa|Result of a number raised to a power
EXP|M|number|e mũ x|e raised to the power of a number
LN|M|number|Lôgarit tự nhiên|Natural logarithm
LOG|M|number, [base]|Lôgarit theo cơ số|Logarithm to a specified base
LOG10|M|number|Lôgarit cơ số 10|Base-10 logarithm
PI|M||Số pi|The value of pi
RAND|M||Số ngẫu nhiên từ 0 đến 1|Random number between 0 and 1
RANDBETWEEN|M|bottom, top|Số nguyên ngẫu nhiên trong khoảng|Random integer between two numbers
CEILING|M|number, significance|Làm tròn lên tới bội số|Rounds up to the nearest multiple
CEILING.MATH|M|number, [significance], [mode]|Làm tròn lên tới bội số|Rounds up to the nearest multiple
FLOOR|M|number, significance|Làm tròn xuống tới bội số|Rounds down to the nearest multiple
FLOOR.MATH|M|number, [significance], [mode]|Làm tròn xuống tới bội số|Rounds down to the nearest multiple
SIGN|M|number|Dấu của số|Sign of a number
FACT|M|number|Giai thừa|Factorial
COMBIN|M|number, number_chosen|Số tổ hợp|Number of combinations
PERMUT|M|number, number_chosen|Số chỉnh hợp|Number of permutations
GCD|M|number1, [number2], ...|Ước chung lớn nhất|Greatest common divisor
LCM|M|number1, [number2], ...|Bội chung nhỏ nhất|Least common multiple
EVEN|M|number|Làm tròn lên số chẵn|Rounds up to the nearest even integer
ODD|M|number|Làm tròn lên số lẻ|Rounds up to the nearest odd integer
SUBTOTAL|M|function_num, ref1, ...|Tổng phụ (bỏ qua hàng bị lọc)|Subtotal in a list (ignores filtered rows)
SIN|M|number|Sin|Sine
COS|M|number|Cos|Cosine
TAN|M|number|Tang|Tangent
ASIN|M|number|Arcsin|Arcsine
ACOS|M|number|Arccos|Arccosine
ATAN|M|number|Arctang|Arctangent
ATAN2|M|x_num, y_num|Arctang theo tọa độ|Arctangent from x and y
SINH|M|number|Sin hyperbol|Hyperbolic sine
COSH|M|number|Cos hyperbol|Hyperbolic cosine
TANH|M|number|Tang hyperbol|Hyperbolic tangent
DEGREES|M|angle|Đổi radian sang độ|Converts radians to degrees
RADIANS|M|angle|Đổi độ sang radian|Converts degrees to radians
ROMAN|M|number|Số La Mã|Roman numeral as text
SEQUENCE|M|rows, [columns], [start], [step]|Dãy số|Sequence of numbers
AVERAGE|S|number1, [number2], ...|Trung bình cộng|Average of the arguments
AVERAGEA|S|value1, [value2], ...|Trung bình (tính cả chữ, logic)|Average including text and logicals
AVERAGEIF|S|range, criteria, [average_range]|Trung bình các ô thỏa điều kiện|Average of cells meeting a condition
AVERAGEIFS|S|average_range, criteria_range1, criteria1, ...|Trung bình thỏa nhiều điều kiện|Average of cells meeting multiple criteria
COUNT|S|value1, [value2], ...|Đếm ô chứa số|Counts cells that contain numbers
COUNTA|S|value1, [value2], ...|Đếm ô không trống|Counts non-empty cells
COUNTBLANK|S|range|Đếm ô trống|Counts empty cells
COUNTIF|S|range, criteria|Đếm ô thỏa điều kiện|Counts cells that meet a condition
COUNTIFS|S|criteria_range1, criteria1, ...|Đếm ô thỏa nhiều điều kiện|Counts cells that meet multiple criteria
MAX|S|number1, [number2], ...|Giá trị lớn nhất|Largest value
MIN|S|number1, [number2], ...|Giá trị nhỏ nhất|Smallest value
MAXA|S|value1, [value2], ...|Lớn nhất (tính cả logic)|Largest value including logicals
MINA|S|value1, [value2], ...|Nhỏ nhất (tính cả logic)|Smallest value including logicals
MAXIFS|S|max_range, criteria_range1, criteria1, ...|Lớn nhất thỏa điều kiện|Maximum among cells meeting criteria
MINIFS|S|min_range, criteria_range1, criteria1, ...|Nhỏ nhất thỏa điều kiện|Minimum among cells meeting criteria
MEDIAN|S|number1, [number2], ...|Trung vị|Median
MODE|S|number1, [number2], ...|Giá trị xuất hiện nhiều nhất|Most frequently occurring value
STDEV|S|number1, [number2], ...|Độ lệch chuẩn mẫu|Sample standard deviation
STDEV.S|S|number1, [number2], ...|Độ lệch chuẩn mẫu|Sample standard deviation
STDEV.P|S|number1, [number2], ...|Độ lệch chuẩn tổng thể|Population standard deviation
VAR|S|number1, [number2], ...|Phương sai mẫu|Sample variance
VAR.S|S|number1, [number2], ...|Phương sai mẫu|Sample variance
VAR.P|S|number1, [number2], ...|Phương sai tổng thể|Population variance
AVEDEV|S|number1, [number2], ...|Độ lệch tuyệt đối trung bình|Average of absolute deviations
GEOMEAN|S|number1, [number2], ...|Trung bình nhân|Geometric mean
HARMEAN|S|number1, [number2], ...|Trung bình điều hòa|Harmonic mean
LARGE|S|array, k|Giá trị lớn thứ k|k-th largest value
SMALL|S|array, k|Giá trị nhỏ thứ k|k-th smallest value
RANK|S|number, ref, [order]|Thứ hạng của số|Rank of a number in a list
RANK.EQ|S|number, ref, [order]|Thứ hạng của số|Rank of a number in a list
RANK.AVG|S|number, ref, [order]|Thứ hạng (trung bình khi bằng)|Rank (average for ties)
PERCENTILE|S|array, k|Phân vị thứ k|k-th percentile
PERCENTILE.INC|S|array, k|Phân vị thứ k|k-th percentile
QUARTILE|S|array, quart|Tứ phân vị|Quartile of a data set
QUARTILE.INC|S|array, quart|Tứ phân vị|Quartile of a data set
CORREL|S|array1, array2|Hệ số tương quan|Correlation coefficient
SLOPE|S|known_y's, known_x's|Độ dốc hồi quy tuyến tính|Slope of the linear regression line
INTERCEPT|S|known_y's, known_x's|Hệ số chặn hồi quy|Intercept of the regression line
FORECAST|S|x, known_y's, known_x's|Dự báo tuyến tính|Linear forecast
IF|L|logical_test, [value_if_true], [value_if_false]|Trả về giá trị theo điều kiện|Returns one value if true and another if false
IFS|L|logical_test1, value1, ...|Kiểm tra nhiều điều kiện|Checks multiple conditions
IFERROR|L|value, value_if_error|Giá trị thay thế khi lỗi|Value if the expression is an error
IFNA|L|value, value_if_na|Giá trị thay thế khi #N/A|Value if the expression is #N/A
AND|L|logical1, [logical2], ...|TRUE nếu tất cả đều đúng|TRUE if all arguments are TRUE
OR|L|logical1, [logical2], ...|TRUE nếu có một đối số đúng|TRUE if any argument is TRUE
NOT|L|logical|Đảo giá trị logic|Reverses the logic
XOR|L|logical1, [logical2], ...|Hoặc loại trừ|Exclusive OR
TRUE|L||Giá trị TRUE|The logical value TRUE
FALSE|L||Giá trị FALSE|The logical value FALSE
SWITCH|L|expression, value1, result1, ..., [default]|Chọn kết quả theo giá trị|Evaluates against a list of values
CHOOSE|L|index_num, value1, [value2], ...|Chọn giá trị theo chỉ số|Chooses a value from a list
VLOOKUP|R|lookup_value, table_array, col_index_num, [range_lookup]|Dò tìm theo cột đầu tiên|Looks up a value in the first column
HLOOKUP|R|lookup_value, table_array, row_index_num, [range_lookup]|Dò tìm theo hàng đầu tiên|Looks up a value in the first row
LOOKUP|R|lookup_value, lookup_vector, [result_vector]|Dò tìm trong vectơ|Looks up a value in a vector
XLOOKUP|R|lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]|Dò tìm linh hoạt|Searches a range and returns a match
INDEX|R|array, row_num, [column_num]|Giá trị tại hàng, cột|Value at a row and column
MATCH|R|lookup_value, lookup_array, [match_type]|Vị trí của giá trị|Relative position of an item
XMATCH|R|lookup_value, lookup_array, [match_mode], [search_mode]|Vị trí của giá trị|Relative position of an item
OFFSET|R|reference, rows, cols, [height], [width]|Tham chiếu dời từ ô gốc|Reference offset from a given reference
INDIRECT|R|ref_text, [a1]|Tham chiếu từ chuỗi|Reference specified by text
ROW|R|[reference]|Số hàng|Row number
ROWS|R|array|Số hàng của vùng|Number of rows
COLUMN|R|[reference]|Số cột|Column number
COLUMNS|R|array|Số cột của vùng|Number of columns
ADDRESS|R|row_num, column_num, [abs_num], [a1], [sheet_text]|Địa chỉ ô dạng chữ|Cell address as text
TRANSPOSE|R|array|Chuyển vị|Transposes an array
HYPERLINK|R|link_location, [friendly_name]|Tạo liên kết|Creates a link
LEFT|T|text, [num_chars]|Ký tự bên trái|Leftmost characters
RIGHT|T|text, [num_chars]|Ký tự bên phải|Rightmost characters
MID|T|text, start_num, num_chars|Ký tự ở giữa|Characters from the middle
LEN|T|text|Độ dài chuỗi|Number of characters
LOWER|T|text|Chữ thường|Converts to lowercase
UPPER|T|text|Chữ HOA|Converts to uppercase
PROPER|T|text|Viết Hoa Đầu Từ|Capitalises each word
TRIM|T|text|Xóa khoảng trắng thừa|Removes extra spaces
CLEAN|T|text|Xóa ký tự không in được|Removes non-printable characters
CONCAT|T|text1, [text2], ...|Nối chuỗi|Joins text
CONCATENATE|T|text1, [text2], ...|Nối chuỗi|Joins text
TEXTJOIN|T|delimiter, ignore_empty, text1, ...|Nối chuỗi với dấu phân cách|Joins text with a delimiter
SUBSTITUTE|T|text, old_text, new_text, [instance_num]|Thay thế chuỗi con|Substitutes new text for old text
REPLACE|T|old_text, start_num, num_chars, new_text|Thay ký tự theo vị trí|Replaces characters by position
FIND|T|find_text, within_text, [start_num]|Tìm (phân biệt hoa thường)|Finds text (case-sensitive)
SEARCH|T|find_text, within_text, [start_num]|Tìm (không phân biệt hoa thường)|Finds text (not case-sensitive)
TEXT|T|value, format_text|Định dạng số thành chữ|Formats a number as text
VALUE|T|text|Chuyển chữ thành số|Converts text to a number
NUMBERVALUE|T|text, [decimal_separator], [group_separator]|Chuyển chữ thành số|Converts text to a number
REPT|T|text, number_times|Lặp chuỗi|Repeats text
EXACT|T|text1, text2|So sánh chính xác|Checks whether two texts are identical
CHAR|T|number|Ký tự theo mã|Character from a code
CODE|T|text|Mã ký tự đầu|Code of the first character
UNICHAR|T|number|Ký tự Unicode|Unicode character
UNICODE|T|text|Mã Unicode|Unicode code point
FIXED|T|number, [decimals], [no_commas]|Định dạng số cố định|Formats a number with fixed decimals
DOLLAR|T|number, [decimals]|Định dạng tiền tệ|Formats as currency text
T|T|value|Trả về chữ|Returns text
TEXTBEFORE|T|text, delimiter, [instance_num]|Chuỗi đứng trước dấu phân cách|Text before a delimiter
TEXTAFTER|T|text, delimiter, [instance_num]|Chuỗi đứng sau dấu phân cách|Text after a delimiter
DOCSO|T|number, [unit]|Đọc số thành chữ tiếng Việt (vd. Một trăm nghìn đồng)|Reads a number in Vietnamese words
DATE|D|year, month, day|Tạo ngày|Creates a date
TIME|D|hour, minute, second|Tạo giờ|Creates a time
TODAY|D||Ngày hôm nay|Today's date
NOW|D||Ngày giờ hiện tại|Current date and time
YEAR|D|serial_number|Năm|Year of a date
MONTH|D|serial_number|Tháng|Month of a date
DAY|D|serial_number|Ngày|Day of a date
HOUR|D|serial_number|Giờ|Hour
MINUTE|D|serial_number|Phút|Minute
SECOND|D|serial_number|Giây|Second
WEEKDAY|D|serial_number, [return_type]|Thứ trong tuần|Day of the week
WEEKNUM|D|serial_number, [return_type]|Số tuần trong năm|Week number
ISOWEEKNUM|D|date|Số tuần ISO|ISO week number
EDATE|D|start_date, months|Ngày sau n tháng|Date a number of months away
EOMONTH|D|start_date, months|Ngày cuối tháng|Last day of the month
DATEDIF|D|start_date, end_date, unit|Khoảng cách giữa hai ngày|Difference between dates
DAYS|D|end_date, start_date|Số ngày giữa hai ngày|Days between two dates
DAYS360|D|start_date, end_date, [method]|Số ngày theo năm 360 ngày|Days based on a 360-day year
NETWORKDAYS|D|start_date, end_date, [holidays]|Số ngày làm việc|Number of working days
WORKDAY|D|start_date, days, [holidays]|Ngày làm việc sau n ngày|Date after a number of working days
DATEVALUE|D|date_text|Chuyển chữ thành ngày|Converts text to a date
TIMEVALUE|D|time_text|Chuyển chữ thành giờ|Converts text to a time
YEARFRAC|D|start_date, end_date, [basis]|Phần năm giữa hai ngày|Fraction of a year
PMT|F|rate, nper, pv, [fv], [type]|Số tiền trả mỗi kỳ|Payment for a loan
FV|F|rate, nper, pmt, [pv], [type]|Giá trị tương lai|Future value
PV|F|rate, nper, pmt, [fv], [type]|Giá trị hiện tại|Present value
NPV|F|rate, value1, [value2], ...|Giá trị hiện tại ròng|Net present value
IRR|F|values, [guess]|Tỷ suất hoàn vốn nội bộ|Internal rate of return
RATE|F|nper, pmt, pv, [fv], [type], [guess]|Lãi suất mỗi kỳ|Interest rate per period
NPER|F|rate, pmt, pv, [fv], [type]|Số kỳ thanh toán|Number of periods
IPMT|F|rate, per, nper, pv, [fv], [type]|Tiền lãi của kỳ|Interest payment for a period
PPMT|F|rate, per, nper, pv, [fv], [type]|Tiền gốc của kỳ|Principal payment for a period
SLN|F|cost, salvage, life|Khấu hao đường thẳng|Straight-line depreciation
SYD|F|cost, salvage, life, per|Khấu hao tổng số năm|Sum-of-years' digits depreciation
DDB|F|cost, salvage, life, period, [factor]|Khấu hao số dư giảm dần kép|Double-declining balance depreciation
ISBLANK|I|value|Ô trống?|TRUE if the value is blank
ISNUMBER|I|value|Là số?|TRUE if the value is a number
ISTEXT|I|value|Là chữ?|TRUE if the value is text
ISNONTEXT|I|value|Không phải chữ?|TRUE if the value is not text
ISLOGICAL|I|value|Là giá trị logic?|TRUE if the value is logical
ISERROR|I|value|Là lỗi?|TRUE if the value is any error
ISERR|I|value|Là lỗi (trừ #N/A)?|TRUE if any error except #N/A
ISNA|I|value|Là #N/A?|TRUE if the value is #N/A
ISEVEN|I|number|Là số chẵn?|TRUE if the number is even
ISODD|I|number|Là số lẻ?|TRUE if the number is odd
ISREF|I|value|Là tham chiếu?|TRUE if the value is a reference
ISFORMULA|I|reference|Ô có công thức?|TRUE if the cell contains a formula
FORMULATEXT|I|reference|Văn bản công thức|Formula as text
NA|I||Giá trị lỗi #N/A|The error value #N/A
TYPE|I|value|Loại giá trị|Type of value
N|I|value|Chuyển thành số|Converts a value to a number
ERROR.TYPE|I|error_val|Mã loại lỗi|Number for an error type
CELL|I|info_type, [reference]|Thông tin ô|Information about a cell
SHEET|I|[value]|Số thứ tự trang tính|Sheet number
SHEETS|I|[reference]|Số trang tính|Number of sheets
";

    private static IEnumerable<string> AllData => [Data, DataDyn, DataFin, DataDist, DataEng, DataMisc];

    private static List<FunctionInfo>? _all;


    public static IReadOnlyList<FunctionInfo> All
    {
        get
        {
            if (_all != null) return _all;
            var list = new List<FunctionInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in AllData.SelectMany(d => d.Split('\n')))
            {
                var p = line.Trim().Split('|');
                if (p.Length < 5) continue;
                if (!seen.Add(p[0])) continue;
                var cat = p[1] switch
                {
                    "M" => FnCategory.Math, "S" => FnCategory.Statistical, "L" => FnCategory.Logical, "R" => FnCategory.Lookup, "T" => FnCategory.Text,
                    "D" => FnCategory.DateTime, "F" => FnCategory.Financial, "E" => FnCategory.Engineering, "B" => FnCategory.Database,
                    "C" => FnCategory.Compatibility, "W" => FnCategory.Web, _ => FnCategory.Information
                };
                list.Add(new FunctionInfo(p[0], cat, p[2], p[3], p[4]));
            }
            return _all = list;
        }
    }

    public static FunctionInfo? Find(string name) => All.FirstOrDefault(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
}
