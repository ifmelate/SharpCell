using System;
using System.Collections.Generic;
using SharpCell.Evaluation;

namespace SharpCell.Functions;

/// <summary>
/// CONVERT with Excel's unit table. Each unit is a factor to its quantity's base unit (gram, meter,
/// second, pascal, newton, joule, watt, tesla, cubic meter, square meter, bit, meter per second);
/// temperatures convert through kelvin with offsets. The factors are the exact definitions where
/// there is one, otherwise the values Excel uses (the atomic mass unit and the electron volt are
/// older CODATA values, the parsec is Excel's own figure).
/// </summary>
internal static class EngineeringConvertFunction
{
    private enum Quantity
    {
        Mass,
        Distance,
        Time,
        Pressure,
        Force,
        Energy,
        Power,
        Magnetism,
        Temperature,
        Volume,
        Area,
        Information,
        Speed,
    }

    private enum Temperature
    {
        None,
        Celsius,
        Fahrenheit,
        Kelvin,
        Rankine,
        Reaumur,
    }

    // Prefixed: whether metric (and, for information, binary) prefixes apply. Dimension: the power
    // a prefix is raised to, as in km2 = (1000 m)^2.
    private sealed record Unit(Quantity Quantity, double Factor, bool Prefixed, int Dimension = 1, Temperature Scale = Temperature.None);

    private const double Inch = 0.0254;
    private const double Foot = 0.3048;
    private const double Yard = 0.9144;
    private const double Mile = 1609.344;
    private const double NauticalMile = 1852;
    private const double LightYear = 9460730472580800;
    private const double PicaPoint = Inch / 72;
    private const double PoundMass = 453.59237;
    private const double Gravity = 9.80665;
    private const double PoundForce = PoundMass / 1000 * Gravity;
    private const double UsGallon = 231 * Inch * Inch * Inch;
    private const double UkGallon = 0.00454609;
    private const double Horsepower = 550 * PoundForce * Foot;

    private static readonly Dictionary<string, Unit> Units = BuildUnits();

    private static readonly Dictionary<string, double> MetricPrefixes = new(StringComparer.Ordinal)
    {
        ["Y"] = 1e24, ["Z"] = 1e21, ["E"] = 1e18, ["P"] = 1e15, ["T"] = 1e12, ["G"] = 1e9, ["M"] = 1e6,
        ["k"] = 1e3, ["h"] = 1e2, ["da"] = 1e1, ["e"] = 1e1, ["d"] = 1e-1, ["c"] = 1e-2, ["m"] = 1e-3,
        ["u"] = 1e-6, ["n"] = 1e-9, ["p"] = 1e-12, ["f"] = 1e-15, ["a"] = 1e-18, ["z"] = 1e-21, ["y"] = 1e-24,
    };

    private static readonly Dictionary<string, double> BinaryPrefixes = new(StringComparer.Ordinal)
    {
        ["Yi"] = Math.Pow(2, 80), ["Zi"] = Math.Pow(2, 70), ["Ei"] = Math.Pow(2, 60), ["Pi"] = Math.Pow(2, 50),
        ["Ti"] = Math.Pow(2, 40), ["Gi"] = Math.Pow(2, 30), ["Mi"] = Math.Pow(2, 20), ["ki"] = Math.Pow(2, 10),
    };

    public static void Register(FunctionRegistry registry) =>
        registry.Add(new FunctionInfo("CONVERT", 3, 3, [ArgumentKind.Value], Convert));

    // Units are case-sensitive text; anything unknown, or two units of different quantities, is #N/A.
    private static Operand Convert(FunctionCall call)
    {
        var number = EngineeringFunctions.StrictNumber(call, 0);
        if (number.IsError)
            return number;
        var fromText = call.Text(1);
        if (fromText.IsError)
            return fromText;
        var toText = call.Text(2);
        if (toText.IsError)
            return toText;

        if (!TryFind(fromText.AsText(), out var from, out var fromScale)
            || !TryFind(toText.AsText(), out var to, out var toScale)
            || from.Quantity != to.Quantity)
            return CellValue.Error(ErrorKind.NA);

        var x = number.AsNumber();
        if (from.Quantity == Quantity.Temperature)
            return CellValue.Number(FromKelvin(ToKelvin(x * fromScale, from.Scale), to.Scale) / toScale);
        return CellValue.Number(x * (from.Factor * fromScale) / (to.Factor * toScale));
    }

    // A unit as written, or a prefix followed by a unit that takes prefixes ("km", "Mibyte",
    // "dam"). Exact names win: "mi" is a mile, not a milli-inch.
    private static bool TryFind(string name, out Unit unit, out double scale)
    {
        scale = 1;
        if (Units.TryGetValue(name, out unit!))
            return true;

        foreach (var (prefixes, informationOnly) in new[] { (BinaryPrefixes, true), (MetricPrefixes, false) })
        {
            foreach (var (prefix, multiplier) in prefixes)
            {
                if (name.Length > prefix.Length && name.StartsWith(prefix, StringComparison.Ordinal)
                    && Units.TryGetValue(name[prefix.Length..], out var baseUnit) && baseUnit.Prefixed
                    && (!informationOnly || baseUnit.Quantity == Quantity.Information))
                {
                    unit = baseUnit;
                    scale = Math.Pow(multiplier, baseUnit.Dimension);
                    return true;
                }
            }
        }

        return false;
    }

    private static double ToKelvin(double x, Temperature scale) => scale switch
    {
        Temperature.Celsius => x + 273.15,
        Temperature.Fahrenheit => (x - 32) / 1.8 + 273.15,
        Temperature.Rankine => x / 1.8,
        Temperature.Reaumur => x * 1.25 + 273.15,
        _ => x,
    };

    private static double FromKelvin(double k, Temperature scale) => scale switch
    {
        Temperature.Celsius => k - 273.15,
        Temperature.Fahrenheit => (k - 273.15) * 1.8 + 32,
        Temperature.Rankine => k * 1.8,
        Temperature.Reaumur => (k - 273.15) * 0.8,
        _ => k,
    };

    private static Dictionary<string, Unit> BuildUnits()
    {
        var units = new Dictionary<string, Unit>(StringComparer.Ordinal);

        void Add(Quantity quantity, double factor, bool prefixed, params string[] names) => AddPowered(quantity, factor, prefixed, 1, names);

        void AddPowered(Quantity quantity, double factor, bool prefixed, int dimension, params string[] names)
        {
            foreach (var name in names)
                units.Add(name, new Unit(quantity, factor, prefixed, dimension));
        }

        // A length unit squared and cubed, written "ft2" or "ft^2".
        void AddAreaAndVolume(string name, double meters, bool prefixed)
        {
            AddPowered(Quantity.Area, meters * meters, prefixed, 2, name + "2", name + "^2");
            AddPowered(Quantity.Volume, meters * meters * meters, prefixed, 3, name + "3", name + "^3");
        }

        void AddTemperature(Temperature scale, bool prefixed, params string[] names)
        {
            foreach (var name in names)
                units.Add(name, new Unit(Quantity.Temperature, 1, prefixed, Scale: scale));
        }

        Add(Quantity.Mass, 1, true, "g");
        Add(Quantity.Mass, PoundForce / Foot * 1000, false, "sg");
        Add(Quantity.Mass, PoundMass, false, "lbm");
        Add(Quantity.Mass, 1.660538782e-24, true, "u");
        Add(Quantity.Mass, PoundMass / 16, false, "ozm");
        Add(Quantity.Mass, 0.06479891, false, "grain");
        Add(Quantity.Mass, PoundMass * 100, false, "cwt", "shweight");
        Add(Quantity.Mass, PoundMass * 112, false, "uk_cwt", "lcwt", "hweight");
        Add(Quantity.Mass, PoundMass * 14, false, "stone");
        Add(Quantity.Mass, PoundMass * 2000, false, "ton");
        Add(Quantity.Mass, PoundMass * 2240, false, "brton", "LTON", "uk_ton");

        Add(Quantity.Distance, 1, true, "m");
        Add(Quantity.Distance, Mile, false, "mi");
        Add(Quantity.Distance, NauticalMile, false, "Nmi");
        Add(Quantity.Distance, Inch, false, "in");
        Add(Quantity.Distance, Foot, false, "ft");
        Add(Quantity.Distance, Yard, false, "yd");
        Add(Quantity.Distance, 1e-10, true, "ang");
        Add(Quantity.Distance, 1.143, false, "ell");
        Add(Quantity.Distance, LightYear, true, "ly");
        Add(Quantity.Distance, 3.0856775812815532e16, true, "parsec", "pc");
        Add(Quantity.Distance, PicaPoint, false, "Picapt", "Pica");
        Add(Quantity.Distance, Inch / 6, false, "pica");
        Add(Quantity.Distance, 6336000.0 / 3937, false, "survey_mi");

        Add(Quantity.Time, 365.25 * 86400, false, "yr");
        Add(Quantity.Time, 86400, false, "day", "d");
        Add(Quantity.Time, 3600, false, "hr");
        Add(Quantity.Time, 60, false, "mn", "min");
        Add(Quantity.Time, 1, true, "sec", "s");

        Add(Quantity.Pressure, 1, true, "Pa", "p");
        Add(Quantity.Pressure, 101325, true, "atm", "at");
        Add(Quantity.Pressure, 133.322, true, "mmHg");
        Add(Quantity.Pressure, PoundForce / (Inch * Inch), false, "psi");
        Add(Quantity.Pressure, 101325.0 / 760, false, "Torr");

        Add(Quantity.Force, 1, true, "N");
        Add(Quantity.Force, 1e-5, true, "dyn", "dy");
        Add(Quantity.Force, PoundForce, false, "lbf");
        Add(Quantity.Force, Gravity / 1000, true, "pond");

        Add(Quantity.Energy, 1, true, "J");
        Add(Quantity.Energy, 1e-7, true, "e");
        Add(Quantity.Energy, 4.184, true, "c");
        Add(Quantity.Energy, 4.1868, true, "cal");
        Add(Quantity.Energy, 1.602176487e-19, true, "eV", "ev");
        Add(Quantity.Energy, Horsepower * 3600, false, "HPh", "hh");
        Add(Quantity.Energy, 3600, true, "Wh", "wh");
        Add(Quantity.Energy, PoundForce * Foot, false, "flb");
        Add(Quantity.Energy, 1055.05585262, false, "BTU", "btu");

        Add(Quantity.Power, Horsepower, false, "HP", "h");
        Add(Quantity.Power, 735.49875, false, "PS");
        Add(Quantity.Power, 1, true, "W", "w");

        Add(Quantity.Magnetism, 1, true, "T");
        Add(Quantity.Magnetism, 1e-4, true, "ga");

        AddTemperature(Temperature.Celsius, false, "C", "cel");
        AddTemperature(Temperature.Fahrenheit, false, "F", "fah");
        AddTemperature(Temperature.Kelvin, true, "K", "kel");
        AddTemperature(Temperature.Rankine, false, "Rank");
        AddTemperature(Temperature.Reaumur, false, "Reau");

        Add(Quantity.Volume, UsGallon / 768, false, "tsp");
        Add(Quantity.Volume, 5e-6, false, "tspm");
        Add(Quantity.Volume, UsGallon / 256, false, "tbs");
        Add(Quantity.Volume, UsGallon / 128, false, "oz");
        Add(Quantity.Volume, UsGallon / 16, false, "cup");
        Add(Quantity.Volume, UsGallon / 8, false, "pt", "us_pt");
        // Excel accepts prefixes on the U.K. pint (but not on the U.S. one).
        Add(Quantity.Volume, UkGallon / 8, true, "uk_pt");
        Add(Quantity.Volume, UsGallon / 4, false, "qt");
        Add(Quantity.Volume, UkGallon / 4, false, "uk_qt");
        Add(Quantity.Volume, UsGallon, false, "gal");
        Add(Quantity.Volume, UkGallon, false, "uk_gal");
        Add(Quantity.Volume, 0.001, true, "l", "L", "lt");
        Add(Quantity.Volume, UsGallon * 42, false, "barrel");
        Add(Quantity.Volume, 0.03523907016688, false, "bushel");
        Add(Quantity.Volume, 100 * Foot * Foot * Foot, false, "GRT", "regton");
        Add(Quantity.Volume, 40 * Foot * Foot * Foot, false, "MTON");

        Add(Quantity.Area, 4046.8564224, false, "uk_acre");
        Add(Quantity.Area, 4046.8726098742522, false, "us_acre");
        Add(Quantity.Area, 100, true, "ar");
        Add(Quantity.Area, 10000, false, "ha");
        Add(Quantity.Area, 2500, false, "Morgen");

        AddAreaAndVolume("m", 1, true);
        AddAreaAndVolume("ang", 1e-10, true);
        AddAreaAndVolume("ft", Foot, false);
        AddAreaAndVolume("in", Inch, false);
        AddAreaAndVolume("yd", Yard, false);
        AddAreaAndVolume("mi", Mile, false);
        AddAreaAndVolume("Nmi", NauticalMile, false);
        AddAreaAndVolume("ly", LightYear, false);
        AddAreaAndVolume("Picapt", PicaPoint, false);
        AddAreaAndVolume("Pica", PicaPoint, false);

        Add(Quantity.Information, 1, true, "bit");
        Add(Quantity.Information, 8, true, "byte");

        Add(Quantity.Speed, 6080 * Foot / 3600, false, "admkn");
        Add(Quantity.Speed, NauticalMile / 3600, false, "kn");
        Add(Quantity.Speed, 1.0 / 3600, true, "m/h", "m/hr");
        Add(Quantity.Speed, 1, true, "m/s", "m/sec");
        // Excel accepts prefixes on miles per hour too.
        Add(Quantity.Speed, Mile / 3600, true, "mph");

        return units;
    }
}
