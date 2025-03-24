from units import NamedComposedUnit, scaled_unit, unit
from units.exception import IncompatibleUnitsError
from units.predefined import define_units
from units.registry import REGISTRY


class UnitConversionError(Exception):
    """Raises when unit conversion goes wrong"""


def define_energy_model_units():
    scaled_unit("%", "percent", 1)

    scaled_unit("kW", "W", 1e3)
    scaled_unit("MW", "kW", 1e3)
    scaled_unit("GW", "MW", 1e3)
    scaled_unit("TW", "GW", 1e3)

    scaled_unit("min", "s", 60)
    scaled_unit("h", "min", 60)
    scaled_unit("day", "h", 24)
    scaled_unit("a", "day", 365)

    scaled_unit("kg", "g", 1e3)
    scaled_unit("t", "kg", 1e3)
    scaled_unit("kt", "t", 1e3)
    scaled_unit("Mt", "kt", 1e3)
    scaled_unit("Gt", "Mt", 1e3)

    scaled_unit("kp", "p", 1e3)
    scaled_unit("Mp", "kp", 1e3)
    scaled_unit("Gp", "Mp", 1e3)

    scaled_unit("100km", "km", 1e2)
    scaled_unit("Mm", "km", 1e3)
    scaled_unit("Gm", "Mm", 1e3)

    scaled_unit("kpkm", "pkm", 1e3)
    scaled_unit("Mpkm", "kpkm", 1e3)
    scaled_unit("Gpkm", "Mpkm", 1e3)
    scaled_unit("Bpkm", "Mpkm", 1e3)

    scaled_unit("ktkm", "tkm", 1e3)
    scaled_unit("Mtkm", "ktkm", 1e3)
    scaled_unit("Gtkm", "Mtkm", 1e3)
    scaled_unit("Btkm", "Mtkm", 1e3)

    scaled_unit("tCO2eq", "kgCO2eq", 1e3)
    scaled_unit("ktCO2eq", "tCO2eq", 1e3)
    scaled_unit("MtCO2eq", "ktCO2eq", 1e3)
    scaled_unit("GtCO2eq", "MtCO2eq", 1e3)

    scaled_unit("kEUR", "EUR", 1e3)
    scaled_unit("MEUR", "kEUR", 1e3)
    scaled_unit("BEUR", "MEUR", 1e3)
    scaled_unit("TEUR", "BEUR", 1e3)
    scaled_unit("M€", "MEUR", 1)

    scaled_unit("k_units", "units", 1e3)
    scaled_unit("M_units", "k_units", 1e3)
    scaled_unit("G_units", "M_units", 1e3)
    scaled_unit("T_units", "G_units", 1e3)
    scaled_unit("Million units", "M_units", 1)

    scaled_unit("vehicle", "vehicle", 1)
    scaled_unit("vehicles", "vehicle", 1)
    scaled_unit("number vehicles", "vehicles", 1)
    scaled_unit("kvehicles", "vehicles", 1e3)
    scaled_unit("Mvehicles", "kvehicles", 1e3)
    scaled_unit("Gvehicles", "Mvehicles", 1e3)
    scaled_unit("Tvehicles", "Gvehicles", 1e3)

    scaled_unit("PJ", "J", 1e15)
    scaled_unit("GJ", "J", 1e9)
    scaled_unit("TJ", "J", 1e12)
    scaled_unit("kWh", "J", 3.6e6)  # 1 kWh = 3.6e6 Joules
    scaled_unit("kWh", "PJ", 3.6e-9)

    NamedComposedUnit("Wh", unit("W") * unit("h"))
    NamedComposedUnit("kWh", unit("kW") * unit("h"))
    NamedComposedUnit("MWh", unit("MW") * unit("h"))
    NamedComposedUnit("GWh", unit("GW") * unit("h"))
    NamedComposedUnit("TWh", unit("TW") * unit("h"))

    NamedComposedUnit("MWh/MWh", unit("MWh") * unit("MWh"))
    NamedComposedUnit("GWh/GWh", unit("GWh") * unit("GWh"))
    NamedComposedUnit("kWh*day", unit("kWh") * unit("day"))
    NamedComposedUnit("MWh*day", unit("MWh") * unit("day"))
    NamedComposedUnit("GWh*day", unit("GWh") * unit("day"))
    NamedComposedUnit("kWh*a", unit("kWh") * unit("a"))
    NamedComposedUnit("MWh*a", unit("MWh") * unit("a"))
    NamedComposedUnit("GWh*a", unit("GWh") * unit("a"))
    NamedComposedUnit("t*a", unit("t") * unit("a"))
    NamedComposedUnit("t*day", unit("t") * unit("day"))
    NamedComposedUnit("kt*day", unit("kt") * unit("day"))
    NamedComposedUnit("Mt*day", unit("Mt") * unit("day"))
    NamedComposedUnit("Gt*day", unit("Gt") * unit("day"))
    NamedComposedUnit("kt*a", unit("kt") * unit("a"))
    NamedComposedUnit("Mt*a", unit("Mt") * unit("a"))
    NamedComposedUnit("Gt*a", unit("Gt") * unit("a"))
    NamedComposedUnit("vehicle*a", unit("vehicle") * unit("a"))
    NamedComposedUnit("vehicle*day", unit("vehicle") * unit("day"))

    NamedComposedUnit("kW/h", unit("kW") / unit("h"))
    NamedComposedUnit("MW/h", unit("MW") / unit("h"))
    NamedComposedUnit("GW/h", unit("GW") / unit("h"))
    NamedComposedUnit("TW/h", unit("TW") / unit("h"))

    NamedComposedUnit("MWh/MW", unit("MWh") / unit("MW"))
    NamedComposedUnit("MWh/t", unit("MWh") / unit("t"))
    NamedComposedUnit("MWh/kt", unit("MWh") / unit("kt"))
    NamedComposedUnit("MWh/k_units", unit("MWh") / unit("k_units"))
    NamedComposedUnit("MWh/M_units", unit("MWh") / unit("M_units"))
    NamedComposedUnit("MWh/G_units", unit("MWh") / unit("G_units"))

    # ------ EUR/**  ------
    NamedComposedUnit("EUR/t", unit("EUR") / unit("t"))
    NamedComposedUnit("€/t", unit("EUR") / unit("t"))
    NamedComposedUnit("kEUR/t", unit("kEUR") / unit("t"))
    NamedComposedUnit("MEUR/t", unit("MEUR") / unit("t"))
    NamedComposedUnit("BEUR/t", unit("BEUR") / unit("t"))

    NamedComposedUnit("EUR/Mt", unit("EUR") / unit("Mt"))
    NamedComposedUnit("kEUR/Mt", unit("kEUR") / unit("Mt"))
    NamedComposedUnit("MEUR/Mt", unit("MEUR") / unit("Mt"))
    NamedComposedUnit("BEUR/Mt", unit("BEUR") / unit("Mt"))

    NamedComposedUnit("EUR/kWh", unit("EUR") / unit("kWh"))
    NamedComposedUnit("EUR/MWh", unit("EUR") / unit("MWh"))
    NamedComposedUnit("€/MWh", unit("EUR") / unit("MWh"))
    NamedComposedUnit("€/MWj", unit("€") / unit("MWj"))

    NamedComposedUnit("kEUR/MWh", unit("kEUR") / unit("MWh"))
    NamedComposedUnit("MEUR/MWh", unit("MEUR") / unit("MWh"))

    NamedComposedUnit("EUR/(MWh*a)", unit("EUR") / unit("MWh*a"))
    NamedComposedUnit("kEUR/(MWh*a)", unit("kEUR") / unit("MWh*a"))
    NamedComposedUnit("MEUR/(MWh*a)", unit("MEUR") / unit("MWh*a"))

    NamedComposedUnit("EUR/(Mt*a)", unit("EUR") / unit("Mt*a"))
    NamedComposedUnit("kEUR/(Mt*a)", unit("kEUR") / unit("Mt*a"))
    NamedComposedUnit("MEUR/(Mt*a)", unit("MEUR") / unit("Mt*a"))

    NamedComposedUnit("EUR/(vehicle*a)", unit("EUR") / unit("vehicle*a"))
    NamedComposedUnit("kEUR/(vehicle*a)", unit("kEUR") / unit("vehicle*a"))
    NamedComposedUnit("MEUR/(vehicle*a)", unit("MEUR") / unit("vehicle*a"))

    NamedComposedUnit("EUR/MW", unit("EUR") / unit("MW"))
    NamedComposedUnit("EUR/W", unit("EUR") / unit("W"))

    NamedComposedUnit("€/MW", unit("EUR") / unit("MW"))

    NamedComposedUnit("EUR/(MW*a)", unit("EUR") / (unit("MW") * unit("a")))
    NamedComposedUnit("EUR/MW*a", unit("EUR") / (unit("MW") * unit("a")))
    NamedComposedUnit("€/MW*a", unit("EUR") / (unit("MW") * unit("a")))
    NamedComposedUnit("EUR/MW/a", unit("EUR") / unit("MW") / unit("a"))
    NamedComposedUnit("EUR/W/a", unit("EUR") / unit("W") / unit("a"))

    NamedComposedUnit("EUR/vehicle", unit("EUR") / unit("vehicle"))
    NamedComposedUnit("kEUR/vehicle", unit("kEUR") / unit("vehicle"))
    NamedComposedUnit("MEUR/vehicle", unit("MEUR") / unit("vehicle"))

    NamedComposedUnit("EUR/kW", unit("EUR") / unit("kW"))
    NamedComposedUnit("EUR/W", unit("EUR") / unit("W"))
    NamedComposedUnit("EUR/W*a", unit("EUR") / unit("W") * unit("a"))
    NamedComposedUnit("EUR/kW*a", unit("EUR") / (unit("kW") * unit("a")))

    NamedComposedUnit("EUR/pkm", unit("EUR") / unit("pkm"))
    NamedComposedUnit("MEUR/pkm", unit("MEUR") / unit("pkm"))
    NamedComposedUnit("EUR/kpkm", unit("EUR") / unit("kpkm"))
    NamedComposedUnit("MEUR/kpkm", unit("MEUR") / unit("kpkm"))
    NamedComposedUnit("EUR/Bpkm", unit("EUR") / unit("Bpkm"))
    NamedComposedUnit("MEUR/Bpkm", unit("MEUR") / unit("Bpkm"))
    NamedComposedUnit("MEUR/GW", unit("MEUR") / unit("GW"))
    NamedComposedUnit("M€/GW", unit("MEUR") / unit("GW"))
    NamedComposedUnit("MEUR/Kt CO2-eq", unit("MEUR") / unit("Kt CO2-eq"))
    NamedComposedUnit("M€/Kt CO2-eq", unit("MEUR") / unit("Kt CO2-eq"))
    NamedComposedUnit("MEUR/Million units", unit("MEUR") / unit("M_units"))
    NamedComposedUnit("M€/Million units", unit("MEUR") / unit("M_units"))
    NamedComposedUnit("MEUR/PJ", unit("MEUR") / unit("PJ"))
    NamedComposedUnit("M€/PJ", unit("MEUR") / unit("PJ"))

    NamedComposedUnit("MEUR/GW", unit("MEUR") / unit("GW"))
    NamedComposedUnit("MEUR/(GW*a)", unit("MEUR") / (unit("GW") * unit("a")))
    NamedComposedUnit("MEUR/GWh", unit("MEUR") / unit("GWh"))
    NamedComposedUnit("Kt/Mt", unit("Kt") / unit("Mt"))
    NamedComposedUnit("EUR/kW*a", unit("EUR") / (unit("kW") * unit("a")))

    NamedComposedUnit("kWh/100km", unit("kWh") / unit("100km"))
    NamedComposedUnit("MWh/100km", unit("MWh") / unit("100km"))

    NamedComposedUnit("kg/MWh", unit("kg") / unit("MWh"))
    NamedComposedUnit("kg/kWh", unit("kg") / unit("kWh"))
    NamedComposedUnit("kg/t", unit("kg") / unit("t"))

    NamedComposedUnit("km/a", unit("km") / unit("a"))
    NamedComposedUnit("km/day", unit("km") / unit("day"))
    NamedComposedUnit("Mm/a", unit("Mm") / unit("a"))

    NamedComposedUnit("percent/h", unit("percent") / unit("h"))
    NamedComposedUnit("percent/day", unit("percent") / unit("day"))
    NamedComposedUnit("percent/a", unit("percent") / unit("a"))
    NamedComposedUnit("percent/vehicle", unit("percent") / unit("vehicle"))
    NamedComposedUnit("percent/kvehicles", unit("percent") / unit("kvehicles"))

    NamedComposedUnit("persons/vehicle", unit("persons") / unit("vehicle"))
    NamedComposedUnit("p/vehicle", unit("p") / unit("vehicle"))
    NamedComposedUnit("kp/vehicle", unit("kp") / unit("vehicle"))
    NamedComposedUnit("Mp/vehicle", unit("Mp") / unit("vehicle"))
    NamedComposedUnit("t/vehicle", unit("t") / unit("vehicle"))
    NamedComposedUnit("kg/vehicle", unit("kg") / unit("vehicle"))

    NamedComposedUnit("%/h", unit("%") / unit("h"))

    NamedComposedUnit("Kt/Million units", unit("Kt") / unit("Million units"))
    NamedComposedUnit("Kt/M_units", unit("Kt") / unit("M_units"))

    NamedComposedUnit("km/PJ", unit("km") / unit("PJ"))
    NamedComposedUnit("kWh/km", unit("kWh") / unit("km"))

    # how to define Kt， Mt，PJ？
    NamedComposedUnit("Kt/PJ", unit("kt") / unit("PJ"))
    NamedComposedUnit("kt/PJ", unit("kt") / unit("PJ"))
    NamedComposedUnit("Mt/PJ", unit("Mt") / unit("PJ"))
    NamedComposedUnit("g/GJ", unit("g") / unit("GJ"))
    NamedComposedUnit("t/TJ", unit("t") / unit("TJ"))

    NamedComposedUnit("PJ/Million units", unit("PJ") / unit("M_units"))
    NamedComposedUnit("PJ/M_units", unit("PJ") / unit("M_units"))
    NamedComposedUnit("PJ/Mt", unit("PJ") / unit("Mt"))
    NamedComposedUnit("PJ/kt", unit("PJ") / unit("kt"))

    NamedComposedUnit("kWh/100km", unit("kWh") / unit("100km"))
    NamedComposedUnit("kWh/km", unit("kWh") / unit("km"))

    NamedComposedUnit("kW/a", unit("kW") / unit("a"))
    NamedComposedUnit("W/a", unit("W") / unit("a"))

    NamedComposedUnit("t/MWh", unit("t") / unit("MWh"))
    NamedComposedUnit("km/(vehicle*a)", unit("km") / (unit("vehicle") * unit("a")))


def get_conversion_factor(convert_from, convert_to):
    # Check if units exist in the registry
    if convert_from not in REGISTRY:
        raise UnitConversionError(f"Unknown from-unit '{convert_from}'.")
    if convert_to not in REGISTRY:
        raise UnitConversionError(f"Unknown to-unit '{convert_to}'.")

    # Handle normal conversions
    try:
        return unit(convert_to)(unit(convert_from)(1)).get_num()
    except IncompatibleUnitsError:
        # Check for reciprocal match (invertible units)
        try:
            # Attempt the reciprocal conversion
            reciprocal_factor = unit(convert_to)(1 / unit(convert_from)(1)).get_num()
            return 1 / reciprocal_factor
        except IncompatibleUnitsError:
            raise UnitConversionError(
                f"Cannot convert from unit '{convert_from}' to unit '{convert_to}'"
            )


def convert_unit(value, from_unit, to_unit):
    try:
        if (from_unit == "kWh/km" and to_unit == "km/PJ") or (
            from_unit == "kWh/100km" and to_unit == "km/PJ"
        ):
            factor = get_conversion_factor(from_unit, to_unit)
            return (1 / (value * factor)), 1
        else:
            factor = get_conversion_factor(from_unit, to_unit)
            return (value * factor), 0

    except UnitConversionError as e:
        print(f"Error in conversion: {e}")
        return None, 0


# Define standard and energy model units
define_units()
define_energy_model_units()
