using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Tests;

/// <summary>
/// Probability distributions and hypothesis tests. Expected values were computed independently with
/// scipy 1.18 (scipy.stats / scipy.special) or are the examples printed in Microsoft's function reference.
/// </summary>
public static class EngineTestsDist
{
    private static Workbook NewData()
    {
        var wb = Workbook.CreateBlank(1);
        var s = wb.Sheets[0];
        double[] a = [3, 4, 5, 8, 9, 1, 2, 4, 5], b = [6, 19, 3, 2, 14, 4, 5, 17, 1], z = [3, 6, 7, 8, 6, 5, 4, 2, 1, 9];
        for (int i = 0; i < a.Length; i++) { s.SetValue(i, 0, CellValue.Number(a[i])); s.SetValue(i, 1, CellValue.Number(b[i])); }
        s.SetValue(9, 0, CellValue.Text("n/a")); // paired T.TEST skips the pair, A1:A10 vs B1:B10 still match in size
        s.SetValue(9, 1, CellValue.Number(7));
        for (int i = 0; i < z.Length; i++) s.SetValue(i, 6, CellValue.Number(z[i]));
        wb.Recalculate();
        return wb;
    }

    private static CellValue Ev(Workbook wb, string f) => wb.Calc.EvaluateFormula(wb.Sheets[0], 20, 9, f);

    public static void Run(TestRunner t)
    {
        var wb = NewData();
        int asserts = 0;
        // relative tolerance (tail probabilities can be tiny, so an absolute tolerance would prove nothing)
        void R(string f, double expected, double rel)
        {
            asserts++;
            var v = Ev(wb, f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > rel * Math.Abs(expected) + 1e-300)
                throw new AssertException($"{f} => {v.Kind}:{(v.IsNumber ? v.Num.ToString("R") : v.ToString())} expected {expected:R}");
        }
        // documented example, rounded to the decimals shown in the documentation
        void D(string f, double expected, int decimals)
        {
            asserts++;
            var v = Ev(wb, f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > 0.5 * Math.Pow(10, -decimals) + 1e-12)
                throw new AssertException($"{f} => {v} expected {expected} ({decimals} decimals)");
        }
        void E(string f, string code) { asserts++; var v = Ev(wb, f); if (!v.IsError || v.Str != code) throw new AssertException($"{f} => {v} expected {code}"); }

        t.Test("Distributions: values (scipy reference + Microsoft examples)", () =>
        {
            R("NORM.DIST(42,40,1.5,TRUE)", 0.9087887802741321, 1e-12);
            R("NORM.DIST(42,40,1.5,FALSE)", 0.10934004978399577, 1e-12);
            R("NORMDIST(35,40,1.5,TRUE)", 0.0004290603331968372, 1e-12);
            R("NORM.S.DIST(1.333333,TRUE)", 0.9087887256040952, 1e-12);
            R("NORM.S.DIST(1.333333,FALSE)", 0.16401014756936722, 1e-12);
            R("NORM.S.DIST(-8,TRUE)", 6.22096057427174e-16, 1e-12);
            R("NORM.S.DIST(-20,TRUE)", 2.7536241186061556e-89, 1e-12);
            R("NORM.S.DIST(0.1,TRUE)", 0.539827837277029, 1e-12);
            R("NORMSDIST(-1.5)", 0.06680720126885807, 1e-12);
            R("NORM.INV(0.908789,40,1.5)", 42.00000200956616, 1e-12);
            R("NORMINV(0.02,10,3)", 3.838753268104531, 1e-12);
            R("NORM.S.INV(0.908789)", 1.3333346730441076, 1e-12);
            R("NORM.S.INV(0.3)", -0.5244005127080409, 1e-12);
            R("NORM.S.INV(1E-10)", -6.361340902404056, 1e-12);
            R("NORM.S.INV(0.999999)", 4.753424308817087, 1e-12);
            R("NORM.S.INV(1E-300)", -37.0470962993612, 1e-12);
            R("NORMSINV(0.975)", 1.959963984540054, 1e-12);
            R("PHI(0.75)", 0.30113743215480443, 1e-12);
            R("PHI(-3)", 0.0044318484119380075, 1e-12);
            R("GAUSS(2)", 0.4772498680518208, 1e-12);
            R("GAUSS(-0.001)", -0.0003989422139110626, 1e-12);
            R("GAUSS(1E-10)", 3.989422804014327e-11, 1e-12);
            R("STANDARDIZE(42,40,1.5)", 1.3333333333333333, 1e-12);
            R("LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.03908355570680047, 1e-12);
            R("LOGNORM.DIST(4,3.5,1.2,FALSE)", 0.01761759668181923, 1e-12);
            R("LOGNORMDIST(30,3.5,1.2)", 0.4671899565868674, 1e-12);
            R("LOGNORM.INV(0.039084,3.5,1.2)", 4.000025218680635, 1e-12);
            R("LOGINV(0.9,0,1)", 3.6022244792791573, 1e-12);
            R("CONFIDENCE.NORM(0.05,2.5,50)", 0.6929519121748389, 1e-12);
            R("CONFIDENCE(0.05,2.5,50.9)", 0.6929519121748389, 1e-12);
            R("CONFIDENCE.T(0.05,1,50)", 0.28419685549572987, 1e-09);
            R("CONFIDENCE.T(0.1,3,4)", 3.530045152202735, 1e-09);
            R("Z.TEST({3,6,7,8,6,5,4,2,1,9},4)", 0.09057419685136386, 1e-12);
            R("Z.TEST({3,6,7,8,6,5,4,2,1,9},6)", 0.8630433891295299, 1e-12);
            R("ZTEST({3,6,7,8,6,5,4,2,1,9},4,2)", 0.04099516050019152, 1e-12);
            R("Z.TEST(G1:G10,4)", 0.09057419685136386, 1e-12);
            R("BINOM.DIST(6,10,0.5,FALSE)", 0.2050781249999999, 1e-12);
            R("BINOM.DIST(6,10,0.5,TRUE)", 0.828125, 1e-12);
            R("BINOMDIST(6.9,10,0.5,TRUE)", 0.828125, 1e-12);
            R("BINOM.DIST(3,1000,0.001,TRUE)", 0.9810731665496396, 1e-12);
            R("BINOM.DIST(400,1000,0.3,FALSE)", 4.011150660289922e-12, 1e-12);
            R("BINOM.DIST(0,20,0.05,FALSE)", 0.35848592240854227, 1e-12);
            R("BINOM.DIST.RANGE(60,0.75,48)", 0.08397496742904753, 1e-12);
            R("BINOM.DIST.RANGE(60,0.75,45,50)", 0.5236297934718879, 1e-12);
            R("BINOM.DIST.RANGE(5000,0.5,1000,3000)", 1.0, 1e-12);
            R("BINOM.INV(6,0.5,0.75)", 4.0, 0);
            R("CRITBINOM(100,0.3,0.9)", 36.0, 0);
            R("BINOM.INV(1000000,0.01,0.05)", 9837.0, 0);
            R("NEGBINOM.DIST(10,5,0.25,TRUE)", 0.3135140584781766, 1e-12);
            R("NEGBINOM.DIST(10,5,0.25,FALSE)", 0.05504866037517788, 1e-12);
            R("NEGBINOMDIST(3,2,0.6)", 0.09215999999999999, 1e-12);
            R("HYPGEOM.DIST(1,4,8,20,TRUE)", 0.4654282765737874, 1e-12);
            R("HYPGEOM.DIST(1,4,8,20,FALSE)", 0.36326109391124867, 1e-12);
            R("HYPGEOMDIST(1,4,8,20)", 0.36326109391124867, 1e-12);
            R("HYPGEOM.DIST(30,100,400,1000,TRUE)", 0.01932340133619305, 1e-12);
            R("HYPGEOM.DIST(50,100,400,1000,TRUE)", 0.9875331366894732, 1e-12);
            R("POISSON.DIST(2,5,TRUE)", 0.12465201948308108, 1e-12);
            R("POISSON.DIST(2,5,FALSE)", 0.08422433748856832, 1e-12);
            R("POISSON(2.7,5,FALSE)", 0.08422433748856832, 1e-12);
            R("POISSON.DIST(150,100,TRUE)", 0.9999987669055809, 1e-12);
            R("POISSON.DIST(1000,1000,FALSE)", 0.012614611348721499, 1e-12);
            R("EXPON.DIST(0.2,10,TRUE)", 0.8646647167633873, 1e-12);
            R("EXPON.DIST(0.2,10,FALSE)", 1.353352832366127, 1e-12);
            R("EXPONDIST(1E-10,1,TRUE)", 9.999999999500001e-11, 1e-12);
            R("WEIBULL.DIST(105,20,100,TRUE)", 0.9295813900692769, 1e-12);
            R("WEIBULL.DIST(105,20,100,FALSE)", 0.035588864024504334, 1e-12);
            R("WEIBULL(2,0.5,3,FALSE)", 0.09021825434110352, 1e-12);
            R("GAMMA.DIST(10.00001131,9,2,FALSE)", 0.03263913041829396, 1e-12);
            R("GAMMA.DIST(10.00001131,9,2,TRUE)", 0.0680940038697872, 1e-12);
            R("GAMMADIST(0.5,0.3,1,TRUE)", 0.8138118046743924, 1e-12);
            R("GAMMADIST(0.5,0.3,1,FALSE)", 0.32936294221710455, 1e-12);
            R("GAMMA.DIST(1000,900,1,TRUE)", 0.9993774022157249, 1e-09);
            R("GAMMA.INV(0.068094,9,2)", 10.000011191437181, 1e-09);
            R("GAMMAINV(0.5,0.3,1)", 0.07313113586695198, 1e-09);
            R("GAMMA.INV(0.999,100,3)", 401.3107917341358, 1e-09);
            R("GAMMA.INV(1E-8,0.5,1)", 7.853981633974481e-17, 1e-09);
            R("BETA.DIST(2,8,10,TRUE,1,3)", 0.6854705810546875, 1e-12);
            R("BETA.DIST(2,8,10,FALSE,1,3)", 1.4837646484374996, 1e-12);
            R("BETA.DIST(0.3,0.5,0.7,TRUE)", 0.45212784899417957, 1e-12);
            R("BETA.DIST(0.3,0.5,0.7,FALSE)", 0.8108944396100788, 1e-12);
            R("BETA.DIST(0.97,200,3,TRUE)", 0.05673453694079385, 1e-09);
            R("BETADIST(2,8,10,1,3)", 0.6854705810546875, 1e-12);
            R("BETA.INV(0.685470581,8,10,1,3)", 1.9999999999631428, 1e-09);
            R("BETAINV(0.1,0.5,0.5)", 0.024471741852423214, 1e-09);
            R("BETA.INV(1E-8,2,3)", 4.082594021609242e-05, 1e-09);
            R("BETA.INV(0.999,30,2)", 0.9985123138630576, 1e-09);
            R("CHISQ.DIST(0.5,1,TRUE)", 0.5204998778130466, 1e-12);
            R("CHISQ.DIST(2,3,FALSE)", 0.2075537487102974, 1e-12);
            R("CHISQ.DIST(2,3.7,FALSE)", 0.2075537487102974, 1e-12);
            R("CHISQ.DIST.RT(18.307,10)", 0.05000058909139812, 1e-12);
            R("CHIDIST(100,10)", 5.4497019829205215e-17, 1e-12);
            R("CHISQ.DIST(1E-6,1,TRUE)", 0.0007978844278221254, 1e-12);
            R("CHISQ.INV(0.93,1)", 3.283020286759539, 1e-09);
            R("CHISQ.INV(0.6,2)", 1.83258146374831, 1e-09);
            R("CHISQ.INV(1E-6,5)", 0.012896160206497087, 1e-09);
            R("CHISQ.INV.RT(0.050001,10)", 18.30697345696106, 1e-09);
            R("CHIINV(1E-12,3)", 58.919755683202155, 1e-09);
            R("CHISQ.INV(0.5,1000)", 999.333412403381, 1e-09);
            R("CHISQ.TEST({58,35;11,25;10,23},{45.35,47.65;17.56,18.44;16.09,16.91})", 0.0003081920170083095, 1e-12);
            R("CHITEST({10,20,30},{20,20,20})", 0.006737946999085468, 1e-12);
            R("F.DIST(15.2069,6,4,TRUE)", 0.9900000430027627, 1e-12);
            R("F.DIST(15.2069,6,4,FALSE)", 0.00122379170878318, 1e-12);
            R("F.DIST(0.5,1,1,FALSE)", 0.30010543871903544, 1e-12);
            R("F.DIST.RT(15.2068649,6,4)", 0.0099999999524646, 1e-12);
            R("FDIST(3,5,30)", 0.025936992113207345, 1e-12);
            R("F.DIST.RT(50,3,200)", 3.707865238759724e-24, 1e-12);
            R("F.INV(0.01,6,4)", 0.10930991412457852, 1e-09);
            R("F.INV.RT(0.01,6,4)", 15.206864861157523, 1e-09);
            R("FINV(0.05,3,1000)", 2.61380358069362, 1e-09);
            R("F.INV(0.99,1,1)", 4052.1806954768217, 1e-09);
            R("F.TEST({6,7,9,15,21},{20,28,31,38,40})", 0.6483178467861745, 1e-12);
            R("FTEST({20,28,31,38,40},{6,7,9,15,21})", 0.6483178467861744, 1e-12);
            R("T.DIST(60,1,TRUE)", 0.9946953263673767, 1e-12);
            R("T.DIST(8,3,FALSE)", 0.0007369065209469264, 1e-12);
            R("T.DIST(-2.5,7,TRUE)", 0.020496109292876444, 1e-12);
            R("T.DIST(0.3,1000000,FALSE)", 0.38138770372344605, 1e-12);
            R("T.DIST(0.001,5,TRUE)", 0.5003796066139012, 1e-12);
            R("T.DIST.2T(1.959999998,60)", 0.054644929975920874, 1e-12);
            R("T.DIST.RT(1.959999998,60)", 0.027322464987960437, 1e-12);
            R("T.DIST.RT(-1,5)", 0.8183912661754386, 1e-12);
            R("T.DIST.RT(30,4)", 3.6764280488306595e-06, 1e-12);
            R("TDIST(1.959999998,60,2)", 0.054644929975920874, 1e-12);
            R("TDIST(1.959999998,60.8,1)", 0.027322464987960437, 1e-12);
            R("T.INV(0.75,2)", 0.8164965809277261, 1e-09);
            R("T.INV(0.01,7)", -2.9979515668685286, 1e-09);
            R("T.INV(1E-10,3)", -2225.769284683093, 1e-09);
            R("T.INV(0.3,1)", -0.7265425280053609, 1e-09);
            R("T.INV(0.9,150)", 1.2872209136149464, 1e-09);
            R("T.INV.2T(0.546449,60)", 0.6065330758257549, 1e-09);
            R("TINV(0.05,10)", 2.2281388519862753, 1e-09);
            R("T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,1)", 0.19601578492528204, 1e-12);
            R("T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},1,2)", 0.09599794338019811, 1e-12);
            R("TTEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1,30},2,3)", 0.10605686248408544, 1e-09);
            R("T.TEST(A1:A9,B1:B9,2,2)", 0.19199588676039622, 1e-12);
            R("T.TEST(A1:A10,B1:B10,1,1)", 0.09800789246264102, 1e-12);
            R("GAMMA(2.5)", 1.329340388179137, 1e-12);
            R("GAMMA(-3.75)", 0.26786612886141653, 1e-12);
            R("GAMMA(5)", 24.0, 1e-12);
            R("GAMMA(170.5)", 5.562092414559999e+305, 1e-12);
            R("GAMMA(1E-5)", 99999.42279422554, 1e-12);
            R("GAMMALN(4)", 1.791759469228055, 1e-12);
            R("GAMMALN.PRECISE(4)", 1.791759469228055, 1e-12);
            R("GAMMALN(0.5)", 0.5723649429247, 1e-12);
            R("GAMMALN(1E-8)", 18.42068073818021, 1e-12);
            R("GAMMALN(1234.5)", 7550.550901077894, 1e-12);
            R("GAMMALN(1E10)", 220258509288.81058, 1e-12);
            R("FISHER(0.75)", 0.9729550745276566, 1e-12);
            R("FISHERINV(0.972955)", 0.7499999673941484, 1e-12);
            R("PERMUTATIONA(3,2)", 9.0, 1e-12);
            R("PERMUTATIONA(2.9,3.1)", 8.0, 1e-12);
            R("PERMUTATIONA(0,0)", 1.0, 1e-12);
            R("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", 0.1, 1e-12);
            R("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)", 0.8, 1e-12);
            // Microsoft documentation examples (values as printed in the docs)
            D("BETA.DIST(2,8,10,TRUE,1,3)", 0.6854706, 7);
            D("BETA.DIST(2,8,10,FALSE,1,3)", 1.4837646, 7);
            D("BETA.INV(0.685470581,8,10,1,3)", 2.0, 6);
            D("BINOM.DIST(6,10,0.5,FALSE)", 0.2050781, 7);
            D("BINOM.DIST.RANGE(60,0.75,48)", 0.084, 3);
            D("BINOM.DIST.RANGE(60,0.75,45,50)", 0.524, 3);
            D("BINOM.INV(6,0.5,0.75)", 4.0, 0);
            D("CHISQ.DIST(0.5,1,TRUE)", 0.52049988, 8);
            D("CHISQ.DIST(2,3,FALSE)", 0.20755375, 8);
            D("CHISQ.DIST.RT(18.307,10)", 0.0500006, 7);
            D("CHISQ.INV(0.93,1)", 3.283020287, 9);
            D("CHISQ.INV(0.6,2)", 1.832581464, 9);
            D("CHISQ.INV.RT(0.050001,10)", 18.30697346, 8);
            D("CHISQ.TEST({58,35;11,25;10,23},{45.35,47.65;17.56,18.44;16.09,16.91})", 0.0003082, 7);
            D("CONFIDENCE.NORM(0.05,2.5,50)", 0.692952, 6);
            D("CONFIDENCE.T(0.05,1,50)", 0.284196855, 9);
            D("EXPON.DIST(0.2,10,TRUE)", 0.86466472, 8);
            D("EXPON.DIST(0.2,10,FALSE)", 1.35335283, 8);
            D("F.DIST(15.2069,6,4,TRUE)", 0.99, 2);
            D("F.DIST(15.2069,6,4,FALSE)", 0.0012238, 7);
            D("F.DIST.RT(15.2068649,6,4)", 0.01, 2);
            D("F.INV(0.01,6,4)", 0.10930991, 8);
            D("F.INV.RT(0.01,6,4)", 15.20686, 5);
            D("F.TEST({6,7,9,15,21},{20,28,31,38,40})", 0.64831785, 8);
            D("FISHER(0.75)", 0.9729551, 7);
            D("FISHERINV(0.972955)", 0.75, 6);
            D("GAMMA(2.5)", 1.329, 3);
            D("GAMMA(-3.75)", 0.268, 3);
            D("GAMMA.DIST(10.00001131,9,2,FALSE)", 0.032639, 6);
            D("GAMMA.DIST(10.00001131,9,2,TRUE)", 0.068094, 6);
            D("GAMMA.INV(0.068094,9,2)", 10.0000112, 7);
            D("GAMMALN(4)", 1.791759, 6);
            D("GAMMALN.PRECISE(4)", 1.791759469, 9);
            D("GAUSS(2)", 0.47725, 5);
            D("HYPGEOM.DIST(1,4,8,20,TRUE)", 0.4654, 4);
            D("HYPGEOM.DIST(1,4,8,20,FALSE)", 0.3633, 4);
            D("LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.0390836, 7);
            D("LOGNORM.DIST(4,3.5,1.2,FALSE)", 0.0176176, 7);
            D("LOGNORM.INV(0.039084,3.5,1.2)", 4.0000252, 7);
            D("NEGBINOM.DIST(10,5,0.25,TRUE)", 0.3135141, 7);
            D("NEGBINOM.DIST(10,5,0.25,FALSE)", 0.0550487, 7);
            D("NORM.DIST(42,40,1.5,TRUE)", 0.9087888, 7);
            D("NORM.DIST(42,40,1.5,FALSE)", 0.10934, 5);
            D("NORM.INV(0.908789,40,1.5)", 42.000002, 6);
            D("NORM.S.DIST(1.333333,TRUE)", 0.908788726, 9);
            D("NORM.S.DIST(1.333333,FALSE)", 0.164010148, 9);
            D("NORM.S.INV(0.908789)", 1.3333347, 7);
            D("PERMUTATIONA(3,2)", 9.0, 0);
            D("PHI(0.75)", 0.301137432, 9);
            D("POISSON.DIST(2,5,TRUE)", 0.124652, 6);
            D("POISSON.DIST(2,5,FALSE)", 0.084224, 6);
            D("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", 0.1, 1);
            D("PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)", 0.8, 1);
            D("STANDARDIZE(42,40,1.5)", 1.33333333, 8);
            D("T.DIST(60,1,TRUE)", 0.99469533, 8);
            D("T.DIST(8,3,FALSE)", 0.00073691, 8);
            D("T.DIST.2T(1.959999998,60)", 0.054645, 6);
            D("T.DIST.RT(1.959999998,60)", 0.027322, 6);
            D("T.INV(0.75,2)", 0.8164966, 7);
            D("T.INV.2T(0.546449,60)", 0.606533, 6);
            D("T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,1)", 0.196016, 6);
            D("WEIBULL.DIST(105,20,100,TRUE)", 0.929581, 6);
            D("WEIBULL.DIST(105,20,100,FALSE)", 0.035589, 6);
            D("Z.TEST({3,6,7,8,6,5,4,2,1,9},4)", 0.090574, 6);
            D("Z.TEST({3,6,7,8,6,5,4,2,1,9},6)", 0.863043, 6);
        });
        t.Test("Distributions: argument errors", () =>
        {
            // normal family
            E("NORM.S.INV(0)", Errors.Num); E("NORM.S.INV(1)", Errors.Num); E("NORM.INV(0.5,0,0)", Errors.Num); E("NORM.DIST(1,0,-1,TRUE)", Errors.Num);
            E("NORM.S.DIST(\"abc\",TRUE)", Errors.Value); E("NORMSDIST(1/0)", Errors.Div0); E("STANDARDIZE(1,2,0)", Errors.Num);
            E("LOGNORM.DIST(0,3.5,1.2,TRUE)", Errors.Num); E("LOGNORMDIST(4,3.5,0)", Errors.Num); E("LOGINV(1,0,1)", Errors.Num);
            E("CONFIDENCE.NORM(1,1,10)", Errors.Num); E("CONFIDENCE(0.05,0,10)", Errors.Num); E("CONFIDENCE.T(0.05,1,1)", Errors.Div0); E("CONFIDENCE.T(0.05,1,0.5)", Errors.Num);
            E("ZTEST({1},1)", Errors.Div0); E("Z.TEST(C1:C5,1)", Errors.NA);
            // discrete
            E("BINOM.DIST(11,10,0.5,TRUE)", Errors.Num); E("BINOMDIST(1,10,1.5,FALSE)", Errors.Num); E("BINOM.DIST(-1,10,0.5,FALSE)", Errors.Num);
            E("BINOM.DIST.RANGE(60,0.75,50,45)", Errors.Num); E("BINOM.DIST.RANGE(60,0.75,61)", Errors.Num);
            E("BINOM.INV(10,0.5,1.5)", Errors.Num); E("CRITBINOM(-1,0.5,0.5)", Errors.Num);
            E("NEGBINOM.DIST(10,0,0.25,FALSE)", Errors.Num); E("NEGBINOMDIST(-1,2,0.5)", Errors.Num);
            E("HYPGEOM.DIST(5,4,8,20,FALSE)", Errors.Num); E("HYPGEOMDIST(1,4,8,0)", Errors.Num); E("HYPGEOM.DIST(0,15,8,20,FALSE)", Errors.Num);
            E("POISSON.DIST(-1,5,TRUE)", Errors.Num); E("POISSON(1,-5,TRUE)", Errors.Num);
            // continuous
            E("EXPON.DIST(-1,1,TRUE)", Errors.Num); E("EXPONDIST(1,0,TRUE)", Errors.Num); E("WEIBULL.DIST(1,0,1,TRUE)", Errors.Num);
            E("GAMMA.DIST(1,0,1,TRUE)", Errors.Num); E("GAMMADIST(-1,1,1,TRUE)", Errors.Num); E("GAMMA.INV(1.5,9,2)", Errors.Num);
            E("BETA.DIST(4,8,10,TRUE,1,3)", Errors.Num); E("BETA.DIST(0.5,0,10,TRUE)", Errors.Num); E("BETA.INV(0,8,10)", Errors.Num); E("BETAINV(0.5,8,10,3,1)", Errors.Num);
            E("CHISQ.DIST(-1,2,TRUE)", Errors.Num); E("CHISQ.DIST(1,0.5,TRUE)", Errors.Num); E("CHIDIST(1,1E11)", Errors.Num);
            E("CHISQ.INV(1.1,2)", Errors.Num); E("CHISQ.INV.RT(0,2)", Errors.Num); E("CHIINV(0.5,0)", Errors.Num);
            E("CHISQ.TEST({1,2,3},{1;2;3})", Errors.NA); E("CHITEST({1,2},{0,2})", Errors.Div0);
            E("F.DIST(1,0,4,TRUE)", Errors.Num); E("F.DIST.RT(-1,2,4)", Errors.Num); E("FDIST(1,2,1E10)", Errors.Num);
            E("F.INV(1.5,6,4)", Errors.Num); E("F.INV.RT(0,6,4)", Errors.Num); E("FINV(0.5,0,4)", Errors.Num);
            E("F.TEST({1},{1,2})", Errors.Div0); E("FTEST({1,1,1},{1,2})", Errors.Div0);
            E("T.DIST(1,0,TRUE)", Errors.Num); E("T.DIST.2T(-1,5)", Errors.Num); E("T.DIST.RT(1,0.9)", Errors.Num); E("TDIST(1,5,3)", Errors.Num); E("TDIST(-1,5,1)", Errors.Num);
            E("T.INV(0,5)", Errors.Num); E("T.INV(1,5)", Errors.Num); E("T.INV.2T(1.5,5)", Errors.Num); E("TINV(0.05,0)", Errors.Num);
            E("T.TEST(A1:A9,B1:B9,3,1)", Errors.Num); E("T.TEST(A1:A9,B1:B9,2,4)", Errors.Num); E("TTEST(A1:A9,B1:B8,2,1)", Errors.NA); E("T.TEST({1},{2,3},2,2)", Errors.Div0);
            // gamma function, Fisher, misc
            E("GAMMA(0)", Errors.Num); E("GAMMA(-2)", Errors.Num); E("GAMMA(172)", Errors.Num); E("GAMMALN(0)", Errors.Num); E("GAMMALN.PRECISE(-1)", Errors.Num);
            E("FISHER(1)", Errors.Num); E("FISHER(-1.5)", Errors.Num); E("FISHERINV(\"x\")", Errors.Value);
            E("PERMUTATIONA(-1,2)", Errors.Num); E("PERMUTATIONA(0,1)", Errors.Num);
            E("PROB({0,1},{0.5,0.6},0)", Errors.Num); E("PROB({0,1,2},{0.5,0.5},0)", Errors.NA); E("PROB({0,1},{0,1},0)", Errors.Num);
        });
        t.Test("Distributions: catalog", () =>
        {
            foreach (var name in new[] { "NORM.DIST", "T.TEST", "BINOM.DIST.RANGE", "NORMSINV", "TTEST", "WEIBULL" })
            {
                asserts++;
                var info = FunctionCatalog.Find(name) ?? throw new AssertException("catalog entry missing: " + name);
                if (!Functions.Table.ContainsKey(name)) throw new AssertException("function not registered: " + name);
                var expectedCat = name.Contains('.') || name == "PHI" ? FnCategory.Statistical : FnCategory.Compatibility;
                if (info.Category != expectedCat) throw new AssertException($"{name}: category {info.Category}");
            }
            asserts++;
            var missing = FunctionCatalog.All.Where(f => f.Category is FnCategory.Statistical or FnCategory.Compatibility && !Functions.Table.ContainsKey(f.Name)).Select(f => f.Name).ToList();
            if (missing.Count > 0) throw new AssertException("catalogued but not implemented: " + string.Join(", ", missing));
        });
        t.Log($"      distribution assertions executed: {asserts}");
    }
}
