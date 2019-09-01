using System;

namespace ExcelBuilderDSL.Excel
{
    internal class Util
    {
        /// <summary>
        /// Get column name from number, 1 based 
        /// </summary>
        /// <param name="columnNumber">Numeric value of excel column</param>
        /// <returns></returns>
        /// <exception>
        ///     ArgumentOutOfRangeException:
        ///         index 0 input
        /// </exception>
        internal static string GetCharColumn(uint columnNumber)
        {

            if(columnNumber ==0)
                throw new ArgumentOutOfRangeException("Index 0 is not a valid value; Index number should be greater than 0.");

            const int minAscIIChar = 65;// AscII values goes from 65 to 90, for uppper case chars
            const int interval = 26; // interval is 25 but need +1 for work over 25 values

            var fistColumnValue =0;
            var lastColumnValue = columnNumber;
            while(interval < lastColumnValue){
                fistColumnValue++;
                lastColumnValue = lastColumnValue - interval ;
            }

            Func<uint,char> getLetters =  (uint c) => ((char)(minAscIIChar + c -1) );

            var letter = "";
                if(columnNumber <= interval || fistColumnValue == 0)
                    letter += getLetters(columnNumber);
                else 
                    letter += GetCharColumn((uint)fistColumnValue) + getLetters(lastColumnValue);

            return letter;
        }
    }
}
