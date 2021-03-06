/**
* Copyright(C) G1ANT Ltd, All rights reserved
* Solution G1ANT.Addon.MSOfficeV2, Project G1ANT.Addon.MSOfficeV2
* www.g1ant.com
*
* Licensed under the G1ANT license.
* See License.txt file in the project root for full license information.
*/


using G1ANT.Language;



namespace G1ANT.Addon.MSOfficeV2
{
    [Command(Name = "outlook.close", Tooltip = "This command allows to close active outlook pogram. It must be initiated at the end of a process.", NeedsDelay = true, IsUnderConstruction = false)]
    public class OutlookCloseCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

        }
        public OutlookCloseCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.BooleanStructure(false));
            var outlookManager = OutlookManager.CurrentOutlook;
            outlookManager.Close();
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.BooleanStructure(true));
        }
    }
}
