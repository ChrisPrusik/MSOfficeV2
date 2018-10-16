/**
* Copyright(C) G1ANT Ltd, All rights reserved
* Solution G1ANT.Addon.MSOfficeV2, Project G1ANT.Addon.MSOfficeV2
* www.g1ant.com
*
* Licensed under the G1ANT license.
* See License.txt file in the project root for full license information.
*/


using G1ANT.Language;


using System;

namespace G1ANT.Addon.MSOfficeV2
{
    [Command(Name = "powerpoint.inserttext", Tooltip = "This command inserts text into current slide.", NeedsDelay = true, IsUnderConstruction = false)]

    public class PowerPointInsertTextCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "text to be placed into current slide")]
            public TextStructure TextToBeInserted { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = false, Tooltip = "position of text")]
            public IntegerStructure TextPosition { get; set; } = new IntegerStructure(1);
        }
        public PowerPointInsertTextCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            PowerPointWrapper ppWrapper = PowerPointManager.CurrentPowerPoint;
            string textToBeInserted = arguments.TextToBeInserted.Value;
            int textPosition = arguments.TextPosition.Value;

            try
            {
                ppWrapper.InsertText(textToBeInserted,textPosition);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to insert text. Message: {ex.Message}", ex);
            }
        }
    }
}
