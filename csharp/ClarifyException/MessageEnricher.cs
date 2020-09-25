using System;
using System.Collections.Generic;

namespace codingdojo
{

    public interface IErrorMessage // Chain of Responsibility
    {
        bool AppliesTo(Exception e);
        string CreateMessage(string formulaName, Exception e);
    }

    public class InvalidExpressionErrorMessage : IErrorMessage
    {
        public bool AppliesTo(Exception e)
        {
            return e.GetType() == typeof(ExpressionParseException);
        }

        public string CreateMessage(string formulaName, Exception e)
        {
            return "Invalid expression found in tax formula [" + formulaName +
                   "]. Check that separators and delimiters use the English locale.";
        }
    }

    public class CircularReferenceErrorMessage : IErrorMessage
    {
        public bool AppliesTo(Exception e)
        {
            return e.GetType() == typeof(SpreadsheetException) &&
                e.Message.StartsWith("Circular Reference");
        }

        public string CreateMessage(string formulaName, Exception e)
        {
            var spreadSheetException = (SpreadsheetException)e;
            return "Circular Reference in spreadsheet related to formula '" + formulaName +
                    "'. Cells: " + spreadSheetException.Cells;
        }
    }

    public class MissingLookupTableErrorMessage : IErrorMessage
    {
        public bool AppliesTo(Exception e)
        {
            return "Object reference not set to an instance of an object".Equals(e.Message) &&
                StackTraceContains(e, "VLookup");
        }

        private bool StackTraceContains(Exception e, string message)
        {
            foreach (var ste in e.StackTrace.Split('\n'))
            {
                if (ste.Contains(message))
                {
                    return true;
                }
            }
            return false;
        }

        public string CreateMessage(string formulaName, Exception e)
        {
            return "Missing Lookup Table";
        }
    }

    public class NoMatchesErrorMessage : IErrorMessage
    {
        public bool AppliesTo(Exception e)
        {
            return e.GetType() == typeof(SpreadsheetException) &&
                "No matches found".Equals(e.Message);
        }

        public string CreateMessage(string formulaName, Exception e)
        {
            var spreadSheetException = (SpreadsheetException)e;
            return "No match found for token [" + spreadSheetException.Token +
                    "] related to formula '" + formulaName + "'.";
        }
    }

    public class GenericErrorMessage : IErrorMessage
    {
        public bool AppliesTo(Exception e)
        {
            return true;
        }

        public string CreateMessage(string formulaName, Exception e)
        {
            return e.Message;
        }
    }

    public class MessageEnricher
    {
        public ErrorResult EnrichError(SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var enrichers = new List<IErrorMessage> { //
                new InvalidExpressionErrorMessage(), //
                new CircularReferenceErrorMessage(), //
                new MissingLookupTableErrorMessage(), //
                new NoMatchesErrorMessage()
            };
            foreach (var enricher in enrichers)
            {
                if (enricher.AppliesTo(e))
                {
                    return CreateErrorResult(enricher, spreadsheetWorkbook, e);
                }
            }

            var finalEnricher = new GenericErrorMessage();
            return CreateErrorResult(finalEnricher, spreadsheetWorkbook, e);
        }

        private ErrorResult CreateErrorResult(IErrorMessage enricher, SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var formulaName = spreadsheetWorkbook.GetFormulaName();
            var presentation = spreadsheetWorkbook.GetPresentation();
            var error = enricher.CreateMessage(formulaName, e);
            return new ErrorResult(formulaName, error, presentation);
        }
    }
}