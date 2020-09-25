using System;

namespace codingdojo
{

    public interface SingleEnricher
    {
        bool Applies(Exception e);
        string ErrorMessage(string formulaName, Exception e);
    }

    public class InvalidExpressionEnricher : SingleEnricher
    {
        public bool Applies(Exception e)
        {
            return e.GetType() == typeof(ExpressionParseException);
        }

        public string ErrorMessage(string formulaName, Exception e)
        {
            return "Invalid expression found in tax formula [" + formulaName +
                   "]. Check that separators and delimiters use the English locale.";;
        }
    }

    public class MessageEnricher
    {
        public ErrorResult EnrichError(SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var formulaName = spreadsheetWorkbook.GetFormulaName();
            var presentation = spreadsheetWorkbook.GetPresentation();

            var ie = new InvalidExpressionEnricher();
            if (ie.Applies(e))
            {
                var error = ie.ErrorMessage(formulaName, e);
                return new ErrorResult(formulaName, error, presentation);
            }

            if (e.Message.StartsWith("Circular Reference"))
            {
                var error = parseCircularReferenceException(e, formulaName);
                return new ErrorResult(formulaName, error, presentation);
            }

            if ("Object reference not set to an instance of an object".Equals(e.Message) && StackTraceContains(e, "VLookup")) {
                var error = "Missing Lookup Table";
                return new ErrorResult(formulaName, error, presentation);
            }

            if ("No matches found".Equals(e.Message))
            {
                var error = parseNoMatchException(e, formulaName);
                return new ErrorResult(formulaName, error, presentation);
            }

            if (true){
                var error = e.Message;
                return new ErrorResult(formulaName, error, presentation);
            }
        }

        private bool StackTraceContains(Exception e, string message)
        {
            foreach (var ste in e.StackTrace.Split('\n'))
            {
                 if (ste.Contains(message))
                    return true;
            }
            return false;
        }

        private string parseNoMatchException(Exception e, string formulaName)
        {
            if (e.GetType() == typeof(SpreadsheetException))
            {
                var we = (SpreadsheetException) e;
                return "No match found for token [" + we.Token+ "] related to formula '" + formulaName + "'.";
            }

            return e.Message;
        }

        private string parseCircularReferenceException(Exception e, string formulaName)
        {
            if (e.GetType() == typeof(SpreadsheetException))
            {
                var we = (SpreadsheetException) e;
                return "Circular Reference in spreadsheet related to formula '" + formulaName + "'. Cells: " +
                       we.Cells;
            }

            return e.Message;
        }
    }
}