using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookAddInMama
{
    public class ReplacementParser
    {

        Queue<char> queue;
        Dictionary<string, string> replacementDictionary;

        /// <summary>
        /// Allows to replace named references with their replacements. The names have to be natural numbers for now.
        /// References in the text to be replaced are indicated by a preceeding $ and might enclose the reference number in {}.
        /// Example: Dictionary("1" -> "you") "Hello $1" and "Hello ${1}" convert to "Hello you",
        /// whereas "Hello $10" converts to "Hello $10" and "Hello ${1}0" converts to "Hello you0"
        /// </summary>
        /// <param name="replacementDictionary">A dictionary the holds the reference name as key and the replacement as value</param>
        public ReplacementParser(Dictionary<string, string> replacementDictionary)
        {
            this.replacementDictionary = replacementDictionary;
        }

       /*Not yet in use
        char indicator;
        char bracketStart;
        char bracketEnd;

        public ReplacementParser()
        { }

        public ReplacementParser(char indicator, char bracketStart, char bracketEnd)
        {
            this.indicator = indicator;
            this.bracketStart = bracketStart;
            this.bracketEnd = bracketEnd;
        }*/
        
        /// <summary>
        /// Does the actual replacement based on the configured dictionary.
        /// </summary>
        /// <param name="input">The text to be parsed</param>
        /// <returns>The text with all possible replacements done</returns>
        public string replaceAll(string input)
        {
            //nothing to replace here
            if (string.IsNullOrEmpty(input) || this.replacementDictionary.Count == 0) return input;

            this.queue = new Queue<char>();

            string output = "";

            //magicQueuing indicates queueing after ${ before }
            bool magicQueuing = false;

            //iterate through input string char by char
            for (int i = 0; i < input.Length; i++)
            {
                char current = input[i];

                switch (current)
                {
                    case '0':
                    case '1':
                    case '2':
                    case '3':
                    case '4':
                    case '5':
                    case '6':
                    case '7':
                    case '8':
                    case '9':
                        //if in queue, store these special characters
                        if (queue.Count > 0) queue.Enqueue(current);
                        else output += current;
                        break;
                    case '{':
                        //if in queue, store these special characters
                        if (queue.Count > 0)
                        {
                            queue.Enqueue(current);
                            //and switch on the magic
                            magicQueuing = true;
                        }
                        else output += current;
                        break;
                    case '}':
                        //queueing ends whenever this character appears
                        if (queue.Count > 0)
                        {
                            queue.Enqueue(current);
                            magicQueuing = false;
                            output += collapseQueue();
                        }
                        else output += current;
                        break;
                    case '$':
                        //queueing ends whenever this character appears
                        if (queue.Count == 1) //only $ in queue
                        {
                            output += collapseQueue();
                        }
                        else if (magicQueuing) //magic queue ends only at }
                        {
                            queue.Enqueue(current);
                        }
                        else
                        {
                            output += collapseQueue();
                            queue.Enqueue(current); //start new queueing
                        }
                        break;
                    default:
                        if (magicQueuing)
                        {
                            queue.Enqueue(current);
                        }
                        else
                        {
                            //queueing ends whenever any other character appears
                            if (queue.Count > 0) output += collapseQueue();
                            output += current;
                        }
                        break;
                }
            }
            //there might be something left. get it!
            output += collapseQueue();
            return output;
        }

        /// <summary>
        /// Collapses the queue and replaces it's content by a replacement, if there is one
        /// </summary>
        /// <returns>the queue content replaced</returns>
        private string collapseQueue()
        {
            if (this.queue.Count == 0) return "";

            string queueString = "";
            while (this.queue.Count > 0)
            {
                queueString += this.queue.Dequeue();
            }

            //as for now the placeholder must be $ followed by a number
            MatchCollection matches = Regex.Matches(queueString, @"^\$([0-9]+)$");
            if (matches.Count == 1)
            {
                //the match is the replacement number
                string match = "";
                if (matches[0].Groups[1].Success) match = matches[0].Groups[1].Value; //$n
                
                //there has not to be a replacement for this placeholder
                if (this.replacementDictionary.ContainsKey(match)) return this.replacementDictionary[match];
            }

            //${n}
            matches = Regex.Matches(queueString, @"^\$\{([a-zA-Z0-9]+)(:.*)?\}$");
            if (matches.Count == 1)
            {
                //the match is the replacement number/name
                string match = "";
                if (matches[0].Groups[1].Success) match = matches[0].Groups[1].Value; //${n}
                
                //there has not to be a replacement for this placeholder
                if (this.replacementDictionary.ContainsKey(match)) return this.replacementDictionary[match];
            }

            return queueString;
        }
    }
}
