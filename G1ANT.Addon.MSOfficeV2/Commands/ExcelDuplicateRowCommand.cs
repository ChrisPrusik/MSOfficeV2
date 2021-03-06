/**
* Copyright(C) G1ANT Ltd, All rights reserved
* Solution G1ANT.Addon.MSOfficeV2, Project G1ANT.Addon.MSOfficeV2
* www.g1ant.com
*
* Licensed under the G1ANT license.
* See License.txt file in the project root for full license information.
*/

using System;

using G1ANT.Language;



namespace G1ANT.Addon.MSOfficeV2
{
    [Command(Name = "excel.duplicaterow", Tooltip = "Duplicates row.")]
    public class ExcelDuplicateRowCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Source row's number")]
            public IntegerStructure Source { get; set; }

            [Argument(Required = true, Tooltip = "Destination row's number")]
            public IntegerStructure Destination { get; set; }
        }
        public ExcelDuplicateRowCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.CurrentExcel.DuplicateRow(arguments.Source.Value, arguments.Destination.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to duplicate row. Source: '{arguments.Source.Value}', destination: '{arguments.Destination.Value}'. Message: {ex.Message}", ex);
            }            
        }
    }
}
