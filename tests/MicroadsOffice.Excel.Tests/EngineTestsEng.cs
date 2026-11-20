using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Tests;

/// <summary>Engineering functions. Expected values come from Microsoft's documented examples, exact arithmetic,
/// or SciPy 1.x / Python cmath (Bessel, erf, complex transcendental functions).</summary>
public static class EngineTestsEng
{
    private static Workbook NewData()
    {
        var wb = Workbook.CreateBlank(1);
        var s = wb.Sheets[0];
        s.SetValue(0, 0, CellValue.Text("1+i"));
        s.SetValue(1, 0, CellValue.Text("2-i"));
        s.SetValue(2, 0, CellValue.Number(3));
        // A4 left blank (skipped by IMSUM/IMPRODUCT)
        s.SetValue(4, 0, CellValue.Text("4j"));
        s.SetValue(0, 1, CellValue.Bool(true));
        wb.Recalculate();
        return wb;
    }

    private static CellValue Ev(Workbook wb, string f) => wb.Calc.EvaluateFormula(wb.Sheets[0], 9, 9, f);

    public static void Run(TestRunner t)
    {
        NumFmt.UseLocale(false);
        var wb = NewData();
        void N(string f, double expected, double tol = 1e-12)
        {
            var v = Ev(wb, f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > tol * Math.Max(1, Math.Abs(expected)))
                throw new AssertException($"{f} => {v.Kind}:{v} expected {expected}");
        }
        void R(string f, double expected, double rel = 1e-10) // relative tolerance (tiny or huge results)
        {
            var v = Ev(wb, f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > rel * Math.Abs(expected))
                throw new AssertException($"{f} => {v.Kind}:{v} expected {expected}");
        }
        void S(string f, string expected)
        {
            var v = Ev(wb, f);
            if (v.ToString() != expected) throw new AssertException($"{f} => {v.Kind}:'{v}' expected '{expected}'");
        }
        void E(string f, string code) { var v = Ev(wb, f); if (!v.IsError || v.Str != code) throw new AssertException($"{f} => {v} expected {code}"); }
        void Cx(string f, double re, double im, double tol = 1e-12) { N($"IMREAL({f})", re, tol); N($"IMAGINARY({f})", im, tol); }

        t.Test("Eng: Bessel functions", () =>
        {
            // Microsoft's examples (shown to 9 decimals)
            N("BESSELJ(1.9,2)", 0.329925728, 1e-9); N("BESSELI(1.5,1)", 0.981666428, 1e-9);
            N("BESSELK(1.5,1)", 0.277387804, 5e-9); N("BESSELY(2.5,1)", 0.145918138, 1e-9);
            // SciPy jv / yv / iv / kv
            N("BESSELJ(1.9,2)", 0.32992572769238715); N("BESSELJ(1.5,0)", 0.5118276717359181); N("BESSELJ(10,5)", -0.2340615281867936);
            N("BESSELJ(30,0)", -0.0863679835810402); N("BESSELJ(40,3)", -0.12614481550582082); R("BESSELJ(30,50)", 2.0581656631563854e-08);
            N("BESSELJ(-2.5,1)", -0.4970941024642741); R("BESSELJ(0.01,10)", 2.6911383392363337e-30); N("BESSELJ(100,30)", 0.08146012958117307);
            N("BESSELJ(1.9,2.9)", 0.32992572769238715); N("BESSELJ(0,0)", 1); N("BESSELJ(0,3)", 0); N("BESSELJ(\"1.5\",0)", 0.5118276717359181);
            N("BESSELY(2.5,1)", 0.14591813796678585); N("BESSELY(0.1,0)", -1.5342386513503667); N("BESSELY(10,0)", 0.05567116728359934);
            N("BESSELY(30,1)", 0.08442557066174726); N("BESSELY(3,5)", -1.9059459538286738); R("BESSELY(0.001,2)", -1273239.8630456673);
            N("BESSELI(1.5,1)", 0.9816664285779078); N("BESSELI(0.5,0)", 1.0634833707413236); R("BESSELI(20,3)", 34592416.34091963);
            R("BESSELI(100,0)", 1.0737517071310736e+42); N("BESSELI(-3,2)", 2.2452124409299516); R("BESSELI(1,10)", 2.7529480398368737e-10);
            N("BESSELK(1.5,1)", 0.2773878004568438); N("BESSELK(0.01,0)", 4.721244730161095); R("BESSELK(5,0)", 0.0036910983340425942);
            R("BESSELK(50,3)", 3.727936773826211e-23); R("BESSELK(0.5,4)", 752.2450979104038);
            E("BESSELJ(1,-1)", Errors.Num); E("BESSELI(1,-1)", Errors.Num); E("BESSELY(0,1)", Errors.Num); E("BESSELY(-1,0)", Errors.Num);
            E("BESSELK(0,1)", Errors.Num); E("BESSELK(-2,1)", Errors.Num); E("BESSELI(1000,0)", Errors.Num);
            E("BESSELJ(\"a\",1)", Errors.Value); E("BESSELJ(TRUE,1)", Errors.Value); E("BESSELK(1,\"x\")", Errors.Value);
        });

        t.Test("Eng: base conversion", () =>
        {
            N("BIN2DEC(1100100)", 100); N("BIN2DEC(1111111111)", -1); N("BIN2DEC(1000000000)", -512); N("BIN2DEC(\"0\")", 0);
            E("BIN2DEC(\"12\")", Errors.Num); E("BIN2DEC(\"11111111111\")", Errors.Num); E("BIN2DEC(-1)", Errors.Num); E("BIN2DEC(TRUE)", Errors.Value);
            S("BIN2HEX(11111011,4)", "00FB"); S("BIN2HEX(1110)", "E"); S("BIN2HEX(1111111111)", "FFFFFFFFFF"); S("BIN2HEX(1111111111,2)", "FFFFFFFFFF");
            E("BIN2HEX(11111011,1)", Errors.Num); E("BIN2HEX(1,-1)", Errors.Num); E("BIN2HEX(1,11)", Errors.Num); E("BIN2HEX(1,\"x\")", Errors.Value);
            S("BIN2OCT(1001,3)", "011"); S("BIN2OCT(1100100)", "144"); S("BIN2OCT(1111111111)", "7777777777"); S("BIN2OCT(1001,3.9)", "011");
            S("DEC2BIN(9,4)", "1001"); S("DEC2BIN(-100)", "1110011100"); S("DEC2BIN(9.9)", "1001"); S("DEC2BIN(\"9\")", "1001");
            S("DEC2BIN(511)", "111111111"); S("DEC2BIN(-512)", "1000000000");
            E("DEC2BIN(512)", Errors.Num); E("DEC2BIN(-513)", Errors.Num); E("DEC2BIN(\"abc\")", Errors.Value); E("DEC2BIN(9,3)", Errors.Num);
            S("DEC2HEX(100,4)", "0064"); S("DEC2HEX(-54)", "FFFFFFFFCA"); S("DEC2HEX(28)", "1C"); S("DEC2HEX(-549755813888)", "8000000000");
            S("DEC2HEX(549755813887)", "7FFFFFFFFF"); E("DEC2HEX(64,1)", Errors.Num); E("DEC2HEX(549755813888)", Errors.Num);
            S("DEC2OCT(58,3)", "072"); S("DEC2OCT(-100)", "7777777634"); S("DEC2OCT(536870911)", "3777777777"); E("DEC2OCT(536870912)", Errors.Num);
            S("HEX2BIN(\"F\",8)", "00001111"); S("HEX2BIN(\"B7\")", "10110111"); S("HEX2BIN(\"FFFFFFFE00\")", "1000000000"); E("HEX2BIN(\"200\")", Errors.Num);
            N("HEX2DEC(\"A5\")", 165); N("HEX2DEC(\"a5\")", 165); N("HEX2DEC(\"FFFFFFFF5B\")", -165); N("HEX2DEC(\"3DA408B9\")", 1034160313);
            N("HEX2DEC(\"8000000000\")", -549755813888); E("HEX2DEC(\"G\")", Errors.Num); E("HEX2DEC(\"FFFFFFFFFFF\")", Errors.Num);
            S("HEX2OCT(\"F\",3)", "017"); S("HEX2OCT(\"3B4E\")", "35516"); S("HEX2OCT(\"FFFFFFFF00\")", "7777777400"); E("HEX2OCT(\"20000000\")", Errors.Num);
            S("OCT2BIN(3,3)", "011"); S("OCT2BIN(7777777000)", "1000000000"); E("OCT2BIN(1000)", Errors.Num); E("OCT2BIN(8)", Errors.Num);
            N("OCT2DEC(54)", 44); N("OCT2DEC(7777777533)", -165); E("OCT2DEC(\"9\")", Errors.Num);
            S("OCT2HEX(100,4)", "0040"); S("OCT2HEX(7777777533)", "FFFFFFFF5B"); E("OCT2HEX(100,1)", Errors.Num);
        });

        t.Test("Eng: bitwise, DELTA, GESTEP", () =>
        {
            N("BITAND(1,5)", 1); N("BITAND(13,25)", 9); N("BITAND(2^48-1,2^47+5)", 140737488355333);
            E("BITAND(-1,1)", Errors.Num); E("BITAND(1.5,1)", Errors.Num); E("BITAND(2^48,1)", Errors.Num); E("BITAND(\"a\",1)", Errors.Value);
            N("BITOR(23,10)", 31); N("BITOR(0,0)", 0); E("BITOR(-2,1)", Errors.Num);
            N("BITXOR(5,3)", 6); N("BITXOR(2^40,1)", 1099511627777); E("BITXOR(1,2^48)", Errors.Num);
            N("BITLSHIFT(4,2)", 16); N("BITLSHIFT(16,-2)", 4); N("BITLSHIFT(1,47)", 140737488355328); E("BITLSHIFT(1,54)", Errors.Num); E("BITLSHIFT(-1,1)", Errors.Num);
            N("BITRSHIFT(13,2)", 3); N("BITRSHIFT(1,-3)", 8); N("BITRSHIFT(5,53)", 0); E("BITRSHIFT(5,-54)", Errors.Num); E("BITRSHIFT(\"x\",1)", Errors.Value);
            N("DELTA(5,4)", 0); N("DELTA(5,5)", 1); N("DELTA(0.5,0)", 0); N("DELTA(0)", 1); N("DELTA(\"2\",2)", 1); E("DELTA(\"a\",1)", Errors.Value);
            N("GESTEP(5,4)", 1); N("GESTEP(5,5)", 1); N("GESTEP(-4,-5)", 1); N("GESTEP(-1)", 0); E("GESTEP(\"a\")", Errors.Value);
        });

        t.Test("Eng: ERF, ERFC", () =>
        {
            N("ERF(0.745)", 0.70792892, 1e-8); N("ERF(1)", 0.84270079, 1e-8); N("ERFC(1)", 0.15729921, 1e-8); // Microsoft's examples
            N("ERF(0.745)", 0.7079289200957376); N("ERF(1)", 0.8427007929497148); N("ERF(-0.5)", -0.5204998778130465);
            N("ERF(3)", 0.9999779095030014); N("ERF(0.1)", 0.1124629160182849); N("ERF(0.5,2)", 0.4748223872059062); N("ERF(0)", 0);
            N("ERF.PRECISE(0.745)", 0.7079289200957376); N("ERF.PRECISE(-0.5)", -0.5204998778130465);
            N("ERFC(1)", 0.15729920705028516); R("ERFC(3)", 2.2090496998585445e-05); R("ERFC(5)", 1.5374597944280347e-12);
            R("ERFC(10)", 2.0884875837625446e-45); N("ERFC(-1)", 1.8427007929497148); N("ERFC(0.1)", 0.8875370839817152);
            N("ERFC.PRECISE(1)", 0.15729920705028516); N("ERFC.PRECISE(-0.5)", 1.5204998778130465);
            E("ERF(\"x\")", Errors.Value); E("ERFC(\"x\")", Errors.Value); E("ERF(1,\"x\")", Errors.Value);
        });

        t.Test("Eng: CONVERT", () =>
        {
            N("CONVERT(1,\"lbm\",\"kg\")", 0.45359237); N("CONVERT(68,\"F\",\"C\")", 20); E("CONVERT(2.5,\"ft\",\"sec\")", Errors.NA);
            N("CONVERT(CONVERT(100,\"ft\",\"m\"),\"ft\",\"m\")", 9.290304); // Microsoft's examples
            N("CONVERT(1,\"mi\",\"km\")", 1.609344); N("CONVERT(1,\"in\",\"cm\")", 2.54); N("CONVERT(1,\"Nmi\",\"m\")", 1852);
            N("CONVERT(1,\"yd\",\"ft\")", 3); N("CONVERT(1,\"ly\",\"m\")", 9460730472580800); N("CONVERT(72,\"Picapt\",\"in\")", 1); N("CONVERT(6,\"pica\",\"in\")", 1);
            N("CONVERT(1,\"stone\",\"lbm\")", 14); N("CONVERT(1,\"ton\",\"lbm\")", 2000); N("CONVERT(1,\"uk_ton\",\"kg\")", 1016.0469088); N("CONVERT(16,\"ozm\",\"lbm\")", 1);
            N("CONVERT(1,\"hr\",\"mn\")", 60); N("CONVERT(1,\"yr\",\"day\")", 365.25); N("CONVERT(1,\"day\",\"s\")", 86400); N("CONVERT(1500,\"ms\",\"sec\")", 1.5);
            N("CONVERT(1,\"atm\",\"mmHg\")", 760); N("CONVERT(1,\"atm\",\"kPa\")", 101.325); N("CONVERT(1,\"Torr\",\"Pa\")", 101325.0 / 760);
            N("CONVERT(1,\"lbf\",\"N\")", 4.4482216152605); N("CONVERT(1,\"N\",\"dyn\")", 100000); N("CONVERT(1,\"kN\",\"N\")", 1000);
            N("CONVERT(1,\"kWh\",\"J\")", 3600000); N("CONVERT(1,\"cal\",\"J\")", 4.1868); N("CONVERT(1,\"BTU\",\"J\")", 1055.05585262); N("CONVERT(1,\"kc\",\"c\")", 1000);
            N("CONVERT(1,\"HP\",\"W\")", 745.69987158227022); N("CONVERT(1,\"PS\",\"W\")", 735.49875); N("CONVERT(1,\"kW\",\"w\")", 1000);
            N("CONVERT(1,\"T\",\"ga\")", 10000); N("CONVERT(1,\"mT\",\"ga\")", 10);
            N("CONVERT(100,\"C\",\"F\")", 212); N("CONVERT(0,\"C\",\"K\")", 273.15); N("CONVERT(0,\"F\",\"C\")", -160.0 / 9); N("CONVERT(100,\"Reau\",\"cel\")", 125);
            N("CONVERT(491.67,\"Rank\",\"F\")", 32); N("CONVERT(1000,\"mK\",\"kel\")", 1); N("CONVERT(-40,\"fah\",\"C\")", -40); N("CONVERT(37,\"C\",\"C\")", 37);
            N("CONVERT(1,\"gal\",\"l\")", 3.785411784); N("CONVERT(1,\"gal\",\"qt\")", 4); N("CONVERT(1,\"cup\",\"oz\")", 8); N("CONVERT(1,\"tbs\",\"tsp\")", 3);
            N("CONVERT(1,\"uk_gal\",\"L\")", 4.54609); N("CONVERT(1,\"m3\",\"l\")", 1000); N("CONVERT(1,\"cm3\",\"ml\")", 1); N("CONVERT(1,\"ft^3\",\"in3\")", 1728);
            N("CONVERT(1,\"barrel\",\"gal\")", 42); N("CONVERT(1,\"GRT\",\"ft3\")", 100);
            N("CONVERT(1,\"ha\",\"m2\")", 10000); N("CONVERT(1,\"km2\",\"m^2\")", 1000000); N("CONVERT(1,\"m2\",\"ft2\")", 1 / 0.09290304);
            N("CONVERT(1,\"uk_acre\",\"yd2\")", 4840); N("CONVERT(1,\"ar\",\"m2\")", 100); N("CONVERT(1,\"Morgen\",\"ar\")", 25);
            N("CONVERT(1,\"byte\",\"bit\")", 8); N("CONVERT(1,\"kibyte\",\"byte\")", 1024); N("CONVERT(1,\"kbyte\",\"byte\")", 1000); N("CONVERT(1,\"Mibit\",\"kibit\")", 1024);
            N("CONVERT(1,\"Gibyte\",\"Mibyte\")", 1024);
            N("CONVERT(100,\"km/h\",\"m/s\")", 1000.0 / 36); N("CONVERT(1,\"mph\",\"km/hr\")", 1.609344); N("CONVERT(1,\"kn\",\"m/h\")", 1852);
            N("CONVERT(1,\"m/sec\",\"m/s\")", 1);
            E("CONVERT(1,\"m\",\"kg\")", Errors.NA); E("CONVERT(1,\"xyz\",\"m\")", Errors.NA); E("CONVERT(1,\"KM\",\"m\")", Errors.NA);
            E("CONVERT(1,\"kft\",\"m\")", Errors.NA); E("CONVERT(1,\"kibit\",\"kim\")", Errors.NA); E("CONVERT(1,\"C\",\"kC\")", Errors.NA);
            E("CONVERT(\"a\",\"m\",\"ft\")", Errors.Value);
        });

        t.Test("Eng: complex numbers (construction, parts, arithmetic)", () =>
        {
            S("COMPLEX(3,4)", "3+4i"); S("COMPLEX(3,4,\"j\")", "3+4j"); S("COMPLEX(0,1)", "i"); S("COMPLEX(1,0)", "1"); // Microsoft's examples
            S("COMPLEX(0,-1)", "-i"); S("COMPLEX(1,-1)", "1-i"); S("COMPLEX(1.5,-2.25)", "1.5-2.25i"); S("COMPLEX(0,0)", "0"); S("COMPLEX(-2,0,\"j\")", "-2");
            S("COMPLEX(0,2.5,\"j\")", "2.5j"); S("COMPLEX(1,1,\"\")", "1+i");
            E("COMPLEX(1,1,\"k\")", Errors.Value); E("COMPLEX(1,1,\"I\")", Errors.Value); E("COMPLEX(\"a\",1)", Errors.Value);
            N("IMABS(\"5+12i\")", 13); N("IMAGINARY(\"3+4i\")", 4); N("IMAGINARY(\"0-j\")", -1); N("IMAGINARY(4)", 0); N("IMREAL(\"6-9i\")", 6);
            N("IMARGUMENT(\"3+4i\")", 0.9272952180016122); S("IMCONJUGATE(\"3+4i\")", "3-4i"); // Microsoft's examples
            N("IMREAL(\"1.5e2-3E-1i\")", 150); N("IMAGINARY(\"1.5e2-3E-1i\")", -0.3); N("IMAGINARY(\"-i\")", -1); N("IMAGINARY(\"+.5j\")", 0.5);
            N("IMREAL(\"7\")", 7); N("IMABS(-3)", 3); S("IMCONJUGATE(\"-j\")", "j"); N("IMARGUMENT(\"-1\")", Math.PI);
            E("IMREAL(\"abc\")", Errors.Num); E("IMREAL(\"3+4I\")", Errors.Num); E("IMREAL(\"3 + 4i\")", Errors.Num); E("IMREAL(\"4i+3\")", Errors.Num);
            E("IMABS(TRUE)", Errors.Value); E("IMARGUMENT(0)", Errors.Div0); E("IMREAL(1/0)", Errors.Div0);
            S("IMDIV(\"-238+240i\",\"10+24i\")", "5+12i"); S("IMPRODUCT(\"3+4i\",\"5-3i\")", "27+11i"); S("IMPRODUCT(\"1+2i\",30)", "30+60i");
            S("IMSUB(\"13+4i\",\"5+3i\")", "8+i"); S("IMSUM(\"3+4i\",\"5-3i\")", "8+i"); // Microsoft's examples
            S("IMSUB(\"1+j\",\"1\")", "j"); S("IMSUB(\"2+i\",\"2+i\")", "0"); S("IMSUM(\"1+j\",2)", "3+j"); S("IMDIV(\"i\",\"2\")", "0.5i");
            S("IMSUM(A1:A3)", "6"); S("IMPRODUCT(A1:A3)", "9+3i"); S("IMSUM(A1:A3,\"-6\",\"-i\")", "-i");
            E("IMSUM(\"1+i\",\"1+j\")", Errors.Value); E("IMSUB(\"i\",\"j\")", Errors.Value); E("IMSUM(A1:A5)", Errors.Value);
            E("IMDIV(\"1+i\",0)", Errors.Num); E("IMPRODUCT(\"1+i\",\"x\")", Errors.Num); E("IMSUM(\"1\",B1)", Errors.Value);
        });

        t.Test("Eng: complex numbers (functions)", () =>
        {
            // Microsoft's examples as text
            S("IMSIN(\"4+3i\")", "-7.61923172032141-6.548120040911i"); S("IMEXP(\"1+i\")", "1.46869393991589+2.28735528717884i");
            S("IMLN(\"3+4i\")", "1.6094379124341+0.927295218001612i"); S("IMSQRT(\"1+i\")", "1.09868411346781+0.455089860562227i");
            S("IMPOWER(\"2+3i\",3)", "-46+9.00000000000001i"); S("IMLOG10(\"3+4i\")", "0.698970004336019+0.402719196273373i");
            S("IMLOG2(\"3+4i\")", "2.32192809488736+1.33780421245098i");
            // Python cmath
            Cx("IMCOS(\"4+3i\")", -6.580663040551157, 7.581552742746545); Cx("IMCOSH(\"4+3i\")", -27.034945603074224, 3.8511533348117775);
            Cx("IMSIN(\"4+3i\")", -7.61923172032141, -6.5481200409110025); Cx("IMSINH(\"4+3i\")", -27.016813258003936, 3.853738037919377);
            Cx("IMTAN(\"4+3i\")", 0.00490825806749606, 1.000709536067233); Cx("IMCOT(\"4+3i\")", 0.004901182394304474, -0.9992669278059017);
            Cx("IMSEC(\"4+3i\")", -0.06529402785794704, -0.07522496030277323); Cx("IMCSC(\"4+3i\")", -0.0754898329158637, 0.0648774713706355);
            Cx("IMSECH(\"4+3i\")", -0.03625349691586887, -0.00516434460775318); Cx("IMCSCH(\"4+3i\")", -0.03627588962862601, -0.005174473184019397);
            Cx("IMEXP(\"1+i\")", 1.4686939399158851, 2.2873552871788423); Cx("IMLN(\"3+4i\")", 1.6094379124341003, 0.9272952180016122);
            Cx("IMLOG10(\"3+4i\")", 0.6989700043360187, 0.4027191962733731); Cx("IMLOG2(\"3+4i\")", 2.321928094887362, 1.3378042124509761);
            Cx("IMSQRT(\"1+i\")", 1.0986841134678098, 0.45508986056222733); Cx("IMSQRT(\"-4\")", 0, 2);
            Cx("IMPOWER(\"2+3i\",-1)", 0.15384615384615385, -0.23076923076923078); Cx("IMPOWER(\"2+3i\",0.5)", 1.6741492280355401, 0.8959774761298381);
            Cx("IMPOWER(\"i\",2)", -1, 0); S("IMCOS(\"j\")", "1.54308063481524"); S("IMEXP(\"0\")", "1"); S("IMSQRT(\"4j\")", "1.4142135623731+1.4142135623731j");
            E("IMLN(0)", Errors.Num); E("IMLOG10(\"0\")", Errors.Num); E("IMCOT(0)", Errors.Num); E("IMCSC(\"0\")", Errors.Num); E("IMCSCH(0)", Errors.Num);
            E("IMPOWER(\"2+3i\",\"x\")", Errors.Value); E("IMSIN(\"x\")", Errors.Num); E("IMEXP(\"1000\")", Errors.Num);
        });

        t.Test("Eng: catalog covers every engineering function", () =>
        {
            var eng = FunctionCatalog.All.Where(f => f.Category == FnCategory.Engineering).ToList();
            T.Eq(54, eng.Count, "engineering catalog entries");
            foreach (var f in eng)
            {
                T.True(Functions.Table.ContainsKey(f.Name), "registered " + f.Name);
                T.True(f.DescVi.Length > 0 && f.DescEn.Length > 0, "descriptions " + f.Name);
            }
        });
    }
}
