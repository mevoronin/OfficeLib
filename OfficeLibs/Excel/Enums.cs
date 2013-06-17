namespace OfficeLibs.Excel
{
    public enum ColorIndexEnum
    {
        None = -4142,
        Auto = -4105,
        LightGray = 15,
        Gray = 16,
        Black = 1,
        Blue1 = 42,
        DarkBlue = 55,
        LightRed = 18,
        Red = 3,
        Yellow = 6,
        LightGreen = 43,
        Green = 14,
        Perlamutr = 33,
        Phiolet = 47
    }
    public enum ColorEnum
    {
        Black = 0,
        Blue = 16711680,
        Cyan = 16776960,
        Green = 65280,
        Magenta = 16711935,
        Red = 255,
        White = 16777215,
        Yellow = 65535
    }

    public enum XlCellType
    {
        xlCellTypeAllFormatConditions = -4172,
        xlCellTypeAllValidation = -4174,
        xlCellTypeBlanks = 4,
        xlCellTypeComments = -4144,
        xlCellTypeConstants = 2,
        xlCellTypeFormulas = -4123,
        xlCellTypeLastCell = 11,
        xlCellTypeSameFormatConditions = -4173,
        xlCellTypeSameValidation = -4175,
        xlCellTypeVisible = 12
    }
    public enum XlDeleteShiftDirection
    {
        xlShiftToLeft = -4159, //Cells are shifted to the left. 
        xlShiftUp = -4162 //Cells are shifted up. 
    }

}