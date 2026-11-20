namespace MicroadsOffice.Excel.Formula;

public static partial class FunctionCatalog
{
    /// <summary>Engineering functions (category E).</summary>
    internal const string DataEng = @"
BESSELI|E|x, n|Hàm Bessel biến dạng In(x)|Returns the modified Bessel function In(x)
BESSELJ|E|x, n|Hàm Bessel Jn(x)|Returns the Bessel function Jn(x)
BESSELK|E|x, n|Hàm Bessel biến dạng Kn(x)|Returns the modified Bessel function Kn(x)
BESSELY|E|x, n|Hàm Bessel Yn(x)|Returns the Bessel function Yn(x)
BIN2DEC|E|number|Đổi số nhị phân sang thập phân|Converts a binary number to decimal
BIN2HEX|E|number, [places]|Đổi số nhị phân sang thập lục phân|Converts a binary number to hexadecimal
BIN2OCT|E|number, [places]|Đổi số nhị phân sang bát phân|Converts a binary number to octal
BITAND|E|number1, number2|Phép AND theo bit của hai số|Returns a bitwise AND of two numbers
BITLSHIFT|E|number, shift_amount|Dịch bit sang trái|Returns a number shifted left by shift_amount bits
BITOR|E|number1, number2|Phép OR theo bit của hai số|Returns a bitwise OR of two numbers
BITRSHIFT|E|number, shift_amount|Dịch bit sang phải|Returns a number shifted right by shift_amount bits
BITXOR|E|number1, number2|Phép XOR theo bit của hai số|Returns a bitwise exclusive OR of two numbers
COMPLEX|E|real_num, i_num, [suffix]|Tạo số phức từ phần thực và phần ảo|Converts real and imaginary coefficients into a complex number
CONVERT|E|number, from_unit, to_unit|Đổi đơn vị đo lường|Converts a number from one measurement system to another
DEC2BIN|E|number, [places]|Đổi số thập phân sang nhị phân|Converts a decimal number to binary
DEC2HEX|E|number, [places]|Đổi số thập phân sang thập lục phân|Converts a decimal number to hexadecimal
DEC2OCT|E|number, [places]|Đổi số thập phân sang bát phân|Converts a decimal number to octal
DELTA|E|number1, [number2]|Kiểm tra hai số có bằng nhau không|Tests whether two values are equal
ERF|E|lower_limit, [upper_limit]|Hàm sai số|Returns the error function
ERF.PRECISE|E|x|Hàm sai số|Returns the error function
ERFC|E|x|Hàm sai số bù|Returns the complementary error function
ERFC.PRECISE|E|x|Hàm sai số bù (tích phân từ x đến vô cực)|Returns the complementary ERF function integrated between x and infinity
GESTEP|E|number, [step]|Kiểm tra số có lớn hơn hoặc bằng ngưỡng không|Tests whether a number is greater than or equal to a threshold value
HEX2BIN|E|number, [places]|Đổi số thập lục phân sang nhị phân|Converts a hexadecimal number to binary
HEX2DEC|E|number|Đổi số thập lục phân sang thập phân|Converts a hexadecimal number to decimal
HEX2OCT|E|number, [places]|Đổi số thập lục phân sang bát phân|Converts a hexadecimal number to octal
IMABS|E|inumber|Mô-đun (giá trị tuyệt đối) của số phức|Returns the absolute value (modulus) of a complex number
IMAGINARY|E|inumber|Phần ảo của số phức|Returns the imaginary coefficient of a complex number
IMARGUMENT|E|inumber|Argument của số phức (góc tính bằng radian)|Returns the argument theta, an angle expressed in radians
IMCONJUGATE|E|inumber|Số phức liên hợp|Returns the complex conjugate of a complex number
IMCOS|E|inumber|Cos của số phức|Returns the cosine of a complex number
IMCOSH|E|inumber|Cos hyperbol của số phức|Returns the hyperbolic cosine of a complex number
IMCOT|E|inumber|Cotang của số phức|Returns the cotangent of a complex number
IMCSC|E|inumber|Cosec của số phức|Returns the cosecant of a complex number
IMCSCH|E|inumber|Cosec hyperbol của số phức|Returns the hyperbolic cosecant of a complex number
IMDIV|E|inumber1, inumber2|Thương của hai số phức|Returns the quotient of two complex numbers
IMEXP|E|inumber|Hàm mũ của số phức|Returns the exponential of a complex number
IMLN|E|inumber|Lôgarit tự nhiên của số phức|Returns the natural logarithm of a complex number
IMLOG10|E|inumber|Lôgarit cơ số 10 của số phức|Returns the base-10 logarithm of a complex number
IMLOG2|E|inumber|Lôgarit cơ số 2 của số phức|Returns the base-2 logarithm of a complex number
IMPOWER|E|inumber, number|Lũy thừa của số phức|Returns a complex number raised to an integer power
IMPRODUCT|E|inumber1, [inumber2], ...|Tích các số phức|Returns the product of complex numbers
IMREAL|E|inumber|Phần thực của số phức|Returns the real coefficient of a complex number
IMSEC|E|inumber|Sec của số phức|Returns the secant of a complex number
IMSECH|E|inumber|Sec hyperbol của số phức|Returns the hyperbolic secant of a complex number
IMSIN|E|inumber|Sin của số phức|Returns the sine of a complex number
IMSINH|E|inumber|Sin hyperbol của số phức|Returns the hyperbolic sine of a complex number
IMSQRT|E|inumber|Căn bậc hai của số phức|Returns the square root of a complex number
IMSUB|E|inumber1, inumber2|Hiệu của hai số phức|Returns the difference between two complex numbers
IMSUM|E|inumber1, [inumber2], ...|Tổng các số phức|Returns the sum of complex numbers
IMTAN|E|inumber|Tang của số phức|Returns the tangent of a complex number
OCT2BIN|E|number, [places]|Đổi số bát phân sang nhị phân|Converts an octal number to binary
OCT2DEC|E|number|Đổi số bát phân sang thập phân|Converts an octal number to decimal
OCT2HEX|E|number, [places]|Đổi số bát phân sang thập lục phân|Converts an octal number to hexadecimal
";
}
