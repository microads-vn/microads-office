namespace MicroadsOffice.Excel.Formula;

public static partial class FunctionCatalog
{
    /// <summary>LET / LAMBDA and dynamic-array functions.</summary>
    internal const string DataDyn = @"
LET|L|name1, value1, [name2, value2], ..., calculation|Đặt tên cho kết quả trung gian rồi dùng lại trong công thức|Assigns names to calculation results
LAMBDA|L|[parameter1, ...], calculation|Tạo hàm tùy chỉnh có tham số|Creates a custom, reusable function
ISOMITTED|I|argument|Tham số LAMBDA có bị bỏ trống?|TRUE if a LAMBDA argument is missing
MAP|L|array1, [array2], ..., lambda|Áp dụng LAMBDA cho từng phần tử của mảng|Applies a LAMBDA to each value of the arrays
REDUCE|L|[initial_value], array, lambda|Gộp mảng thành một giá trị bằng LAMBDA|Reduces an array to an accumulated value
SCAN|L|[initial_value], array, lambda|Giá trị tích lũy qua từng phần tử|Returns each intermediate accumulated value
BYROW|L|array, lambda|Áp dụng LAMBDA cho từng hàng|Applies a LAMBDA to each row
BYCOL|L|array, lambda|Áp dụng LAMBDA cho từng cột|Applies a LAMBDA to each column
MAKEARRAY|L|rows, cols, lambda|Tạo mảng tính bằng LAMBDA theo hàng và cột|Creates an array calculated by a LAMBDA
FILTER|R|array, include, [if_empty]|Lọc dữ liệu theo điều kiện|Filters a range of data based on criteria
SORT|R|array, [sort_index], [sort_order], [by_col]|Sắp xếp dữ liệu|Sorts the contents of a range or array
SORTBY|R|array, by_array1, [sort_order1], ...|Sắp xếp theo giá trị của vùng khác|Sorts by the values in another range
UNIQUE|R|array, [by_col], [exactly_once]|Danh sách giá trị không trùng|Returns the unique values
VSTACK|R|array1, [array2], ...|Ghép các mảng theo chiều dọc|Stacks arrays vertically
HSTACK|R|array1, [array2], ...|Ghép các mảng theo chiều ngang|Stacks arrays horizontally
TAKE|R|array, rows, [columns]|Lấy số hàng/cột ở đầu hoặc cuối|Returns rows or columns from the start or end
DROP|R|array, rows, [columns]|Bỏ số hàng/cột ở đầu hoặc cuối|Excludes rows or columns from the start or end
CHOOSECOLS|R|array, col_num1, [col_num2], ...|Chọn các cột của mảng|Returns the specified columns
CHOOSEROWS|R|array, row_num1, [row_num2], ...|Chọn các hàng của mảng|Returns the specified rows
TOCOL|R|array, [ignore], [scan_by_column]|Chuyển mảng thành một cột|Returns the array as one column
TOROW|R|array, [ignore], [scan_by_column]|Chuyển mảng thành một hàng|Returns the array as one row
WRAPROWS|R|vector, wrap_count, [pad_with]|Chia một dãy thành nhiều hàng|Wraps a row or column into rows
WRAPCOLS|R|vector, wrap_count, [pad_with]|Chia một dãy thành nhiều cột|Wraps a row or column into columns
EXPAND|R|array, rows, [columns], [pad_with]|Mở rộng mảng tới kích thước cho trước|Expands an array to the given size
AREAS|R|reference|Số vùng trong tham chiếu|Number of areas in a reference
TRIMRANGE|R|range, [trim_rows], [trim_cols]|Bỏ hàng/cột trống ở rìa vùng|Excludes empty outer rows and columns
GROUPBY|R|row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array]|Gom nhóm và tổng hợp theo hàng|Groups and aggregates values by rows
CEILING.PRECISE|M|number, [significance]|Làm tròn lên tới bội số (không phụ thuộc dấu)|Rounds up to the nearest multiple, regardless of sign
FLOOR.PRECISE|M|number, [significance]|Làm tròn xuống tới bội số (không phụ thuộc dấu)|Rounds down to the nearest multiple, regardless of sign
ISO.CEILING|M|number, [significance]|Làm tròn lên tới bội số theo chuẩn ISO|Rounds up to the nearest multiple (ISO)
FORECAST.LINEAR|S|x, known_y's, known_x's|Dự báo theo xu hướng tuyến tính|Predicts a value along a linear trend
MODE.SNGL|S|number1, [number2], ...|Giá trị xuất hiện nhiều nhất|Most frequently occurring value
PEARSON|S|array1, array2|Hệ số tương quan Pearson|Pearson correlation coefficient
PERCENTILE.EXC|S|array, k|Phân vị thứ k (k trong khoảng 0..1, không gồm hai đầu)|k-th percentile, k exclusive of 0 and 1
STDEVA|S|value1, [value2], ...|Độ lệch chuẩn mẫu (tính cả chữ và giá trị logic)|Sample standard deviation including text and logicals
STDEVP|C|number1, [number2], ...|Độ lệch chuẩn tổng thể (bản cũ)|Population standard deviation (legacy)
VARP|C|number1, [number2], ...|Phương sai tổng thể (bản cũ)|Population variance (legacy)
VND|T|number, [unit]|Đọc số tiền thành chữ tiếng Việt (giống DOCSO)|Reads an amount in Vietnamese words (same as DOCSO)
PIVOTBY|R|row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order], [filter_array]|Bảng tổng hợp theo hàng và cột|Groups and aggregates by rows and columns
";
}
