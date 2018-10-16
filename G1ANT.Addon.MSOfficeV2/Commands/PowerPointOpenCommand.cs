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
    [Command(Name = "powerpoint.open",Tooltip = "This command opens powerpoint program.", NeedsDelay = true, IsUnderConstruction = false)]
    public class PowerPointOpenCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument]
            public TextStructure Path { get; set; } = new TextStructure(string.Empty);

            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public PowerPointOpenCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            string path = arguments.Path.Value;
            PowerPointWrapper powerPointWrapper = PowerPointManager.AddPowerPoint();
            powerPointWrapper.Open(path);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(powerPointWrapper.Id));
        }
    }
}
