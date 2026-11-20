namespace MicroadsOffice.Excel.Formula;

public static partial class FunctionCatalog
{
    /// <summary>Financial functions added with FunctionsFin.cs (bonds, depreciation, cash-flow schedules, T-bills).</summary>
    internal const string DataFin = @"
ACCRINT|F|issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]|Lãi dồn tích của chứng khoán trả lãi định kỳ|Accrued interest for a security that pays periodic interest
ACCRINTM|F|issue, settlement, rate, [par], [basis]|Lãi dồn tích của chứng khoán trả lãi khi đáo hạn|Accrued interest for a security that pays interest at maturity
AMORDEGRC|F|cost, date_purchased, first_period, salvage, period, rate, [basis]|Khấu hao giảm dần theo hệ số từng kỳ kế toán (chuẩn Pháp)|Depreciation for each accounting period using a depreciation coefficient (French system)
AMORLINC|F|cost, date_purchased, first_period, salvage, period, rate, [basis]|Khấu hao đường thẳng từng kỳ kế toán (chuẩn Pháp)|Depreciation for each accounting period (French system)
COUPDAYBS|F|settlement, maturity, frequency, [basis]|Số ngày từ đầu kỳ trả lãi đến ngày thanh toán|Days from the beginning of the coupon period to the settlement date
COUPDAYS|F|settlement, maturity, frequency, [basis]|Số ngày của kỳ trả lãi chứa ngày thanh toán|Days in the coupon period that contains the settlement date
COUPDAYSNC|F|settlement, maturity, frequency, [basis]|Số ngày từ ngày thanh toán đến ngày trả lãi kế tiếp|Days from the settlement date to the next coupon date
COUPNCD|F|settlement, maturity, frequency, [basis]|Ngày trả lãi kế tiếp sau ngày thanh toán|Next coupon date after the settlement date
COUPNUM|F|settlement, maturity, frequency, [basis]|Số lần trả lãi từ ngày thanh toán đến ngày đáo hạn|Number of coupons payable between settlement and maturity
COUPPCD|F|settlement, maturity, frequency, [basis]|Ngày trả lãi gần nhất trước ngày thanh toán|Previous coupon date before the settlement date
CUMIPMT|F|rate, nper, pv, start_period, end_period, type|Tổng tiền lãi trả giữa hai kỳ|Cumulative interest paid between two periods
CUMPRINC|F|rate, nper, pv, start_period, end_period, type|Tổng tiền gốc trả giữa hai kỳ|Cumulative principal paid between two periods
DB|F|cost, salvage, life, period, [month]|Khấu hao số dư giảm dần cố định|Fixed-declining balance depreciation
DISC|F|settlement, maturity, pr, redemption, [basis]|Tỷ lệ chiết khấu của chứng khoán|Discount rate for a security
DOLLARDE|F|fractional_dollar, fraction|Đổi giá dạng phân số sang số thập phân|Converts a price expressed as a fraction into a decimal number
DOLLARFR|F|decimal_dollar, fraction|Đổi giá dạng thập phân sang phân số|Converts a decimal price into a price expressed as a fraction
DURATION|F|settlement, maturity, coupon, yld, frequency, [basis]|Thời gian đáo hạn bình quân (Macaulay) của trái phiếu|Macaulay duration of a security that pays periodic interest
EFFECT|F|nominal_rate, npery|Lãi suất thực tế hằng năm|Effective annual interest rate
FVSCHEDULE|F|principal, schedule|Giá trị tương lai theo chuỗi lãi suất kép|Future value of a principal after a series of compound interest rates
INTRATE|F|settlement, maturity, investment, redemption, [basis]|Lãi suất của chứng khoán đầu tư toàn bộ|Interest rate for a fully invested security
ISPMT|F|rate, per, nper, pv|Tiền lãi của một kỳ (khoản vay trả gốc đều)|Interest paid during a period of a loan with even principal payments
MDURATION|F|settlement, maturity, coupon, yld, frequency, [basis]|Thời gian đáo hạn bình quân điều chỉnh của trái phiếu|Modified Macaulay duration of a security
MIRR|F|values, finance_rate, reinvest_rate|Tỷ suất hoàn vốn nội bộ điều chỉnh|Modified internal rate of return
NOMINAL|F|effect_rate, npery|Lãi suất danh nghĩa hằng năm|Nominal annual interest rate
ODDFPRICE|F|settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis]|Giá trái phiếu có kỳ trả lãi đầu lẻ (mệnh giá 100)|Price per $100 face value of a security with an odd first period
ODDFYIELD|F|settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis]|Lợi suất trái phiếu có kỳ trả lãi đầu lẻ|Yield of a security with an odd first period
ODDLPRICE|F|settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis]|Giá trái phiếu có kỳ trả lãi cuối lẻ (mệnh giá 100)|Price per $100 face value of a security with an odd last period
ODDLYIELD|F|settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis]|Lợi suất trái phiếu có kỳ trả lãi cuối lẻ|Yield of a security with an odd last period
PDURATION|F|rate, pv, fv|Số kỳ để khoản đầu tư đạt giá trị mong muốn|Number of periods for an investment to reach a specified value
PRICE|F|settlement, maturity, rate, yld, redemption, frequency, [basis]|Giá trái phiếu trả lãi định kỳ (mệnh giá 100)|Price per $100 face value of a security that pays periodic interest
PRICEDISC|F|settlement, maturity, discount, redemption, [basis]|Giá chứng khoán chiết khấu (mệnh giá 100)|Price per $100 face value of a discounted security
PRICEMAT|F|settlement, maturity, issue, rate, yld, [basis]|Giá chứng khoán trả lãi khi đáo hạn (mệnh giá 100)|Price per $100 face value of a security that pays interest at maturity
RECEIVED|F|settlement, maturity, investment, discount, [basis]|Số tiền nhận được khi đáo hạn|Amount received at maturity for a fully invested security
RRI|F|nper, pv, fv|Lãi suất tương đương cho mức tăng trưởng của khoản đầu tư|Equivalent interest rate for the growth of an investment
TBILLEQ|F|settlement, maturity, discount|Lợi suất tương đương trái phiếu của tín phiếu kho bạc|Bond-equivalent yield for a Treasury bill
TBILLPRICE|F|settlement, maturity, discount|Giá tín phiếu kho bạc (mệnh giá 100)|Price per $100 face value for a Treasury bill
TBILLYIELD|F|settlement, maturity, pr|Lợi suất của tín phiếu kho bạc|Yield for a Treasury bill
VDB|F|cost, salvage, life, start_period, end_period, [factor], [no_switch]|Khấu hao số dư giảm dần cho một khoảng kỳ bất kỳ|Depreciation for any period using a declining balance method
XIRR|F|values, dates, [guess]|Tỷ suất hoàn vốn nội bộ của dòng tiền không định kỳ|Internal rate of return for a schedule of cash flows that is not necessarily periodic
XNPV|F|rate, values, dates|Giá trị hiện tại ròng của dòng tiền không định kỳ|Net present value for a schedule of cash flows that is not necessarily periodic
YIELD|F|settlement, maturity, rate, pr, redemption, frequency, [basis]|Lợi suất trái phiếu trả lãi định kỳ|Yield on a security that pays periodic interest
YIELDDISC|F|settlement, maturity, pr, redemption, [basis]|Lợi suất hằng năm của chứng khoán chiết khấu|Annual yield for a discounted security
YIELDMAT|F|settlement, maturity, issue, rate, pr, [basis]|Lợi suất hằng năm của chứng khoán trả lãi khi đáo hạn|Annual yield of a security that pays interest at maturity
";
}
