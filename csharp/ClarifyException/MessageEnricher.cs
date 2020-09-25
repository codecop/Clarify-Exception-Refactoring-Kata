﻿using System;
using System.Collections.Generic;

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
                   "]. Check that separators and delimiters use the English locale.";
        }
    }

    public class CircularReferenceEnricher : SingleEnricher
    {
        public bool Applies(Exception e)
        {
            return e.Message.StartsWith("Circular Reference");
        }

        public string ErrorMessage(string formulaName, Exception e)
        {
            if (e.GetType() == typeof(SpreadsheetException))
            {
                var we = (SpreadsheetException)e;
                return "Circular Reference in spreadsheet related to formula '" + formulaName + "'. Cells: " +
                       we.Cells;
            }

            return e.Message;
        }
    }

    public class MissingLookupTableEnricher : SingleEnricher
    {
        public bool Applies(Exception e)
        {
            return "Object reference not set to an instance of an object".Equals(e.Message) &&
                StackTraceContains(e, "VLookup");
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

        public string ErrorMessage(string formulaName, Exception e)
        {
            return "Missing Lookup Table";
        }
    }

    public class MessageEnricher
    {
        public ErrorResult EnrichError(SpreadsheetWorkbook spreadsheetWorkbook, Exception e)
        {
            var formulaName = spreadsheetWorkbook.GetFormulaName();
            var presentation = spreadsheetWorkbook.GetPresentation();

            var enrichers = new List<SingleEnricher> { //
                new InvalidExpressionEnricher(), //
                new CircularReferenceEnricher(), //
                new MissingLookupTableEnricher(), //
            };
            foreach (var enricher in enrichers)
            {
                if (enricher.Applies(e))
                {
                    var error = enricher.ErrorMessage(formulaName, e);
                    return new ErrorResult(formulaName, error, presentation);
                }
            }

            if ("No matches found".Equals(e.Message))
            {
                var error = parseNoMatchException(e, formulaName);
                return new ErrorResult(formulaName, error, presentation);
            }

            if (true)
            {
                var error = e.Message;
                return new ErrorResult(formulaName, error, presentation);
            }
        }

        private string parseNoMatchException(Exception e, string formulaName)
        {
            if (e.GetType() == typeof(SpreadsheetException))
            {
                var we = (SpreadsheetException)e;
                return "No match found for token [" + we.Token + "] related to formula '" + formulaName + "'.";
            }

            return e.Message;
        }

    }
}