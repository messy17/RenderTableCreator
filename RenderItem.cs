using System;


namespace RenderTableCreator
{
    internal class RenderItem : IComparable<RenderItem>
    {
        public string ImageName { get; set; }
        public string Description { get; set; }
        public int LineNumber { get; set; }
        public int RefCount { get; set; } // keeps track of the number of instances 
        

        public RenderItem(string _imageName, string _description, int lineNumber)
        {
            ImageName = _imageName;
            Description = _description;
            LineNumber = lineNumber;
            RefCount = 1; 
        }

        public int CompareTo(RenderItem other)
        {
            int result = 0;
            int padCount = 0;
            String thisName = this.ImageName;
            String otherName = other.ImageName;

            // Normalize lengths of string
            if (thisName.Length < otherName.Length)
            {
                padCount = otherName.Length - thisName.Length;
                thisName = thisName.PadRight(thisName.Length + padCount);
            }
            else if (thisName.Length > otherName.Length)
            {
                padCount = thisName.Length - otherName.Length;
                otherName = otherName.PadRight(otherName.Length + padCount);
            }


            for (int i = 0; i < thisName.Length; i++)
            {
                //result = CompareChar(thisName[i], otherName[i]);
                result = CompareChar2(ref i, thisName, otherName);

                if (result == 0)
                {
                    continue;
                }
                else
                    return result;

            }

            return result;
        }

        private static int CompareChar(char first, char second)
        {
            // priority
            // space ' '
            // underscore '_' 
            // 0-9 [numbers are sorted by numeric value, not string value]
            // a-z [letters are sorted as normal]                      
            //
            // Critical Errors return -255 
            

            char f = first.ToString().ToLower()[0];
            char s = second.ToString().ToLower()[0];

            // Space
            if ((f == ' ') && (s == ' '))
                return 0;
            else if ((f == ' ') && (s != ' '))
                return -1;
            else if (((f != ' ') && (s == ' ')))
                return 1;
            // Underscore 
            else if ((f == '_') && (s == '_'))
                return 0;
            else if ((f == '_') && (s != '_'))
                return -1;
            else if (((f != '_') && (s == '_')))
                return 1;

            // Numbers 
            bool firstIsDigit = int.TryParse(first.ToString(), out int firstDigit);
            bool secondIsDigit = int.TryParse(second.ToString(), out int secondDigit);

            if (firstIsDigit && secondIsDigit)
            {
                if (firstDigit < secondDigit)
                    return -1;
                if (firstDigit > secondDigit)
                    return 1;
                else
                    return 0;
            }
            else if (firstIsDigit)
                return -1;
            else if (secondIsDigit)
                return 1;

            // Neither inputs are a digit; check for letters (lower chase)
            bool firstIsLetter = char.IsLetter(f);
            bool secondIsLetter = char.IsLetter(s);

            if (firstIsLetter && secondIsLetter)
            {
                if (f < s)
                    return -1;
                if (f > s)
                    return 1;
                else
                    return 0;
            }
            else if (firstIsLetter)
            {
                if (s == '_')
                    return -1;
                else
                    // for now just return -1;
                    // later return -255 for crtical errors that include special characters
                    return -1;
            }
            else if (secondIsLetter)
            {
                if (first == '_')
                    return 1;
                else
                    // later return -255 for crtical errors that include special characters
                    return 1;
            }

            // Neither inputs are underscores, spaces, digits or letters 
            // Let the natural char sorting work its magic.            
            if (f < s)
                return -1;
            else if (f > s)
                return 1;
            else
                return 0;

        }

        private static int CompareNumbers(ref int index, string thisName, string otherName)
        {
            String thisNumString = String.Empty;
            String otherNumString = String.Empty;
            int idx;

            for(idx = index; idx < thisName.Length; idx++)
            {
                bool thisIsDigit = char.IsDigit(thisName[idx]);
                bool otherIsDigit = char.IsDigit(otherName[idx]);

                if (thisIsDigit)
                    thisNumString += thisName[idx];

                if (otherIsDigit)
                    otherNumString += otherName[idx];

                if (!thisIsDigit && !otherIsDigit)
                    break; 
            }

            int.TryParse(thisNumString, out int thisNumber);
            int.TryParse(otherNumString, out int otherNumber);

            index = idx - 1;

            if (thisNumber < otherNumber)
                return -1;
            if (thisNumber > otherNumber)
                return 1;
            else
                return 0; 

        }

        private static int CompareChar2(ref int index, String thisName, String otherName)
        {
            // priority
            // space ' '
            // underscore '_' 
            // 0-9 [numbers are sorted by numeric value, not string value]
            // a-z [letters are sorted as normal]                      
            //
            // Critical Errors return -255 
            
            char f = thisName.ToString().ToLower()[index];
            char s = otherName.ToString().ToLower()[index];

            // Space
            if ((f == ' ') && (s == ' '))
                return 0;
            else if ((f == ' ') && (s != ' '))
                return -1;
            else if (((f != ' ') && (s == ' ')))
                return 1;
            
            // Underscore 
            else if ((f == '_') && (s == '_'))
                return 0;
            else if ((f == '_') && (s != '_'))
                return -1;
            else if (((f != '_') && (s == '_')))
                return 1;

            // Numbers            
            if (char.IsDigit(f) || char.IsDigit(s))
                return CompareNumbers(ref index, thisName, otherName);


            // Neither inputs are a digit; check for letters (lower chase)
            bool firstIsLetter = char.IsLetter(f);
            bool secondIsLetter = char.IsLetter(s);

            if (firstIsLetter && secondIsLetter)
            {
                if (f < s)
                    return -1;
                if (f > s)
                    return 1;
                else
                    return 0;
            }
            //else if (firstIsLetter)
            //{
            //    if (s == '_')
            //        return -1;
            //    else
            //        // for now just return -1;
            //        // later return -255 for crtical errors that include special characters
            //        return -1;
            //}
            //else if (secondIsLetter)
            //{
            //    if (f == '_')
            //        return 1;
            //    else
            //        // later return -255 for crtical errors that include special characters
            //        return 1;
            //}

            // Neither inputs are underscores, spaces, digits or letters 
            // Let the natural char sorting work its magic.            
            if (f < s)
                return -1;
            else if (f > s)
                return 1;
            else
                return 0;

        }
    }
}
