using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ExcelDna.Integration; //https://excel-dna.net/


namespace MiscExcelTools
{
    public class ExcelFunctions
    {

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Interp1D
        [ExcelFunction(Category = "MiscExcelTools", Description = "1-Dimensional Interpolation", IsMacroType = true)]
        public static object Interp1D([ExcelArgument(Description = "Interpolation Method:\n0 = Linear; 1 = Akima; 2 = Nearest-neighbor", Name = "method")] int method,
                                        [ExcelArgument(Description = "x Values", Name = "xValues")] double[] xValues,
                                        [ExcelArgument(Description = "y Values", Name = "yValues")] double[] yValues,
                                        [ExcelArgument(Description = "x value to interpolate on", Name = "xStar")] double xStar)
        {
            // error handling
            // ------------------------------------------------------------------------------------------------
            //check the method
            if (method != 0 && method != 1 && method != 2)
            {
                return "#ERROR: invalid method";
                //return ExcelError.ExcelErrorValue;
            }

            // check the counts
            if (xValues.Count() != yValues.Count())
            {
                return "#ERROR";
            }

            // check for duplicates
            List<double> xDuplicates = new List<double> { };
            for (int i = 0; i < xValues.Count(); i++)
            {
                if (xDuplicates.Contains(xValues[i]))
                {
                    return "#ERROR: duplicates in the set of xValues";
                }
                else
                {
                    xDuplicates.Add(xValues[i]);
                }
            }

            // special cases
            // ------------------------------------------------------------------------------------------------
            if (xValues.Count() == 1)
            {
                return yValues[0];
            }

            int xInd = xValues.ToList().IndexOf(xStar);
            if (xInd != -1)
            {
                return yValues[xInd];
            }

            // ------------------------------------------------------------------------------------------------
            //sort x and y values according to the x values
            Array.Sort(yValues.ToArray(), xValues);
            Array.Sort(xValues.ToArray(), xValues);

            int last = xValues.Count() - 1;
            double xMin = xValues[0];
            double xMax = xValues[last];

            // linear
            // ------------------------------------------------------------------------------------------------
            if (method == 0)
            {
                int ind = 0;
                if (xStar < xMin)
                {
                    ind = 1;
                }
                else if (xStar > xMax)
                {
                    ind = last;
                }
                else
                {
                    for (int i = 0; i <= last; i++)
                    {
                        if (xValues[i] >= xStar)
                        {
                            ind = i;
                            break;
                        }
                    }
                }

                double m = (yValues[ind] - yValues[ind - 1]) / (xValues[ind] - xValues[ind - 1]);
                double interp = yValues[ind] + (xStar - xValues[ind]) * m;
                return interp;
            }//linear

            //akima
            // ------------------------------------------------------------------------------------------------
            if (method == 1)
            {
                alglib.spline1dinterpolant akima_interp;
                alglib.spline1dbuildakima(xValues, yValues, out akima_interp);
                //alglib.spline1dbuildcubic(xValues, yValues, out akima_interp);
                double interp = alglib.spline1dcalc(akima_interp, xStar);
                return interp;
            }//akima

            // nearest-neighbor
            // ------------------------------------------------------------------------------------------------
            if (method == 2)
            {
                int ind = 0;
                if (xStar < xMin)
                {
                    ind = 0;
                }
                else if (xStar > xMax)
                {
                    ind = last;
                }
                else
                {
                    double min_dist = xStar - xValues[0];
                    ind = 0;
                    for (int i = 1; i <= last; i++)
                    {
                        double dist = Math.Abs(xValues[i] - xStar);
                        if (dist < min_dist)
                        {
                            min_dist = dist;
                            ind = i;
                        }
                    }
                }

                double interp = yValues[ind];
                return interp;
            }//nearest-neighbor

            return 999999999999.99;

        }//Interp1D

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Interp2D
        [ExcelFunction(Category = "MiscExcelTools", Description = "2-Dimensional Interpolation", IsMacroType = true)]
        public static object Interp2D([ExcelArgument(Description = "Interpolateion Method:\n0 = Linear; 1 = Akima; 2 = Nearest-neighbor", Name = "method")] int method,
                                        [ExcelArgument(Description = "x0 Values (rows)", Name = "x0Values")] double[] x0Values,
                                        [ExcelArgument(Description = "x1 Values (columns)", Name = "x1Values")] double[] x1Values,
                                        [ExcelArgument(Description = "y Values", Name = "yValues")] double[,] yValues,
                                        [ExcelArgument(Description = "x0 value to interpolate on", Name = "x0Star")] double x0Star,
                                        [ExcelArgument(Description = "x1 value to interpolate on", Name = "x1Star")] double x1Star,
                                        [ExcelArgument(Description = "Akima Dimension:\n0 = rows,1 = columns", Name = "dimension")] int dimension)
        {

            int last0 = x0Values.Count() - 1;
            double x0Min = x0Values[0];
            double x0Max = x0Values[last0];

            int last1 = x1Values.Count() - 1;
            double x1Min = x1Values[0];
            double x1Max = x1Values[last1];

            double xStar = 0;
            double xMin = 0;
            double xMax = 0;
            int last = 0;
            int lastO = 0;
            double[] xValues = { };
            double[] xOValues = { };
            double xOStar = 0;

            if (dimension == 1)
            {
                xStar = x1Star;
                xMin = x1Min;
                xMax = x1Max;
                last = last1;
                lastO = last0;
                xValues = x1Values;
                xOValues = x0Values;
                xOStar = x0Star;
            }
            else
            {
                xStar = x0Star;
                xMin = x0Min;
                xMax = x0Max;
                last = last0;
                lastO = last1;
                xValues = x0Values;
                xOValues = x1Values;
                xOStar = x1Star;
            }

            // (ind is the low end)
            int ind = 0;
            if (xStar <= xMin)
            {
                ind = 0;
            }
            else if (xStar >= xMax)
            {
                ind = last - 1;
            }
            else
            {
                for (int i = 0; i <= last; i++)
                {
                    if (xValues[i] >= xStar)
                    {
                        ind = i - 1;
                        break;
                    }
                }
            }

            //the y values for the x points surrounding xStar
            List<double> y_xlow = new List<double> { };
            List<double> y_xhigh = new List<double> { };
            for (int i = 0; i <= lastO; i++)
            {
                if (dimension == 0)
                {
                    y_xlow.Add(yValues[ind, i]);
                    y_xhigh.Add(yValues[ind + 1, i]);
                }
                else
                {
                    y_xlow.Add(yValues[i, ind]);
                    y_xhigh.Add(yValues[i, ind + 1]);
                }

            }

            //the interpolated y values for the x points surrounding xStar, for xOStar
            double y_xOlow = Convert.ToDouble(Interp1D(method, xOValues, y_xlow.ToArray(), xOStar));
            double y_xOhigh = Convert.ToDouble(Interp1D(method, xOValues, y_xhigh.ToArray(), xOStar));

            //the two x points surrounding xStar, and their two interpolated y values for xOStar
            double[] x = new double[] { xValues[ind], xValues[ind + 1] };
            double[] y = new double[] { y_xOlow, y_xOhigh };

            //we'll do a linear interpolation on these points if the method is linear or akima
            if (method == 0 || method == 1)
            {
                return Convert.ToDouble(Interp1D(0, x, y, xStar));
            }
            //otherwise find the nearest neighbor
            else
            {
                if (xStar - x[0] <= x[1] - xOStar)
                {
                    return y[0];
                }
                else
                {
                    return y[1];
                }
            }

        }//Interp2D

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //TableLookup
        [ExcelFunction(Category = "MiscExcelTools", Description = "Table Lookup by names - returns a value", IsMacroType = true)]
        public static object TableLookup([ExcelArgument(Description = "Row Values", Name = "rowValues")] object[] rValues,
                                        [ExcelArgument(Description = "Column Values", Name = "columnValues")] object[] cValues,
                                        [ExcelArgument(Description = "Table", Name = "table")] object[,] table,
                                        [ExcelArgument(Description = "Row Value to Lookup", Name = "rowValue")] object rValue,
                                        [ExcelArgument(Description = "Column Value to Lookup", Name = "columnValue")] object cValue)
        {
            int rind = rValues.ToList().IndexOf(rValue);
            int cind = cValues.ToList().IndexOf(cValue);
            return table[rind, cind];
        }//TableLookup

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //TableLookup2
        [ExcelFunction(Category = "MiscExcelTools", Description = "Table Lookup by names - returns a table", IsMacroType = true)]
        public static object[,] TableLookup2([ExcelArgument(Description = "Row Values", Name = "rowValues")] object[] rValues,
                                        [ExcelArgument(Description = "Column Values", Name = "columnValues")] object[] cValues,
                                        [ExcelArgument(Description = "Table", Name = "table")] object[,] table,
                                        [ExcelArgument(Description = "Row Values to Lookup", Name = "rowLValues")] object[] rLValue,
                                        [ExcelArgument(Description = "Column Values to Lookup", Name = "columnLValues")] object[] cLValue
            //,[ExcelArgument(Description = "Value on Error", Name = "errValue")] object errValue
            )
        {
            object[,] ret = new object[rLValue.Count(), cLValue.Count()];

            for (int r = 0; r < rLValue.Count(); r++)
            {
                for (int c = 0; c < cLValue.Count(); c++)
                {

                    int rind = rValues.ToList().IndexOf(rLValue[r]);
                    int cind = cValues.ToList().IndexOf(cLValue[c]);
                    if (rind == -1 || cind == -1)
                    {
                        ret[r, c] = "#VALUE!";
                    }
                    else
                    {
                        ret[r, c] = table[rind, cind];
                    }
                }
            }
            return ret;

        }//TableLookup2

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Tabularize
        [ExcelFunction(Category = "MiscExcelTools", Description = "Turn a range into a NxM table", IsMacroType = true)]
        public static object Tabularize([ExcelArgument(Description = "Range", Name = "range")] object[] range,
                                        [ExcelArgument(Description = "Number of rows", Name = "rows")] int rows,
                                        [ExcelArgument(Description = "Number of columns", Name = "columns")] int columns)
        {

            object[,] ret = new object[rows, columns];
            int row = 0;
            for (int i = 0; i < range.Count(); i++)
            {

                int col = i % columns;
                if (col == 0 && i != 0)
                    row += 1;

                if (range[i].GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                    ret[row, col] = "";
                else
                    ret[row, col] = range[i];
            }
            return ret;
        }//Tabularize

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //DeTabularize
        [ExcelFunction(Category = "MiscExcelTools", Description = "Turn an NxM table into a range", IsMacroType = true)]
        public static object[,] DeTabularize([ExcelArgument(Description = "Table", Name = "table")] object[,] table,
                                              [ExcelArgument(Description = "0 = return as row; 1 = return as column", Name = "retColumn")] int retColumn)
        {

            int rows = table.GetLength(0);
            int columns = table.GetLength(1);

            int rdims = 1;
            int cdims = rows * columns;

            bool retc = false;
            if (retColumn == 1)
            {
                retc = true;
                cdims = 1;
                rdims = rows * columns;
            }

            object[,] ret = new object[rdims, cdims];

            int row = 0;
            for (int i = 0; i < (rows * columns); i++)
            {

                int col = i % columns;
                if (col == 0 && i != 0)
                    row += 1;

                object item = "";
                if (table[row, col].GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                {
                    //do nothing
                }
                else
                {
                    item = table[row, col];
                }

                if (retc)
                    ret[i, 0] = item;
                else
                    ret[0, i] = item;
            }
            return ret;
        }//DeTabularize

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Replace2
        [ExcelFunction(Category = "MiscExcelTools", Description = "Find and replace text", IsMacroType = true)]
        public static object Replace2([ExcelArgument(Description = "Text", Name = "text")] object text,
                                        [ExcelArgument(Description = "Text to replace", Name = "find")] object find,
                                        [ExcelArgument(Description = "Replacement", Name = "replace")] object replace,
                                        [ExcelArgument(Description = "0 = observe case; 1 = ignore case", Name = "ignoreCase")] int ignoreCase)
        {
            string _text = Convert.ToString(text);
            string _find = Convert.ToString(find);
            string _replace = Convert.ToString(replace);
            int _ignoreCase = Convert.ToInt16(ignoreCase);

            if (text.GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                _text = "";

            if (find.GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                _find = "";

            if (replace.GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                _replace = "";

            if (_ignoreCase == 1)
                return Regex.Replace(_text, _find, _replace, RegexOptions.IgnoreCase);
            else
                return Regex.Replace(_text, _find, _replace);
        }//Replace2

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Join
        [ExcelFunction(Category = "MiscExcelTools", Description = "Join a range", IsMacroType = true)]
        public static object Join([ExcelArgument(Description = "Range", Name = "range")] object[] range,
                                        [ExcelArgument(Description = "Join by", Name = "joiner")] object joiner,
                                        [ExcelArgument(Description = "0 = as is; 1 = add quotes", Name = "addQuotes")] int addQuotes,
                                        [ExcelArgument(Description = "0 = as is; 1 = ignore blanks", Name = "ignoreBlanks")] int ignoreBlanks)
        {
            string _joiner = Convert.ToString(joiner);
            int _addQuotes = Convert.ToInt16(addQuotes);
            int _ignoreBlanks = Convert.ToInt16(ignoreBlanks);

            if (joiner.GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                _joiner = "";

            string result = "";

            int fne = range.Count() - 1; // first non-empty spot
            for (int i = 0; i < range.Count(); i++)
            {
                string item = Convert.ToString(range[i]);
                if (range[i].GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                {
                    item = "";
                    if (_ignoreBlanks == 1)
                        continue;
                }

                if (i < fne)
                    fne = i;

                if (i != 0)
                {
                    if (i != fne)
                        result += joiner;
                }

                if (addQuotes == 1)
                    item = "\"" + item + "\"";
                result += item;

            }
            return result;
        }//Join

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Split
        [ExcelFunction(Category = "MiscExcelTools", Description = "Split text", IsMacroType = true)]
        public static object Split([ExcelArgument(Description = "Text", Name = "text")] object text,
                                        [ExcelArgument(Description = "What to split by", Name = "splitter")] object splitter,
                                        [ExcelArgument(Description = "Which element to return in the split list", Name = "element")] int element)
        {
            string _text = Convert.ToString(text);
            string _splitter = Convert.ToString(splitter);
            int _element = Convert.ToInt16(element);

            if (text.GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                _text = "";

            if (splitter.GetType() == ExcelDna.Integration.ExcelEmpty.Value.GetType())
                _splitter = "";

            return _text.Split(new string[] { _splitter }, StringSplitOptions.None)[_element];

        }//Split

        //-------------------------------------------------------------------------------------------------------------------------------------------------------------------

    }//ExcelFunctions class

}//MiscExcelTools
