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
    [Command(Name = "powerpoint.insertslide", Tooltip = "This command inserts slide into current presentation.", NeedsDelay = true, IsUnderConstruction = false)]

    public class PowerPointInsertSlideCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument]
            public IntegerStructure SlidePosition { get; set; } = new IntegerStructure(1);

            [Argument(Required = false, Tooltip = "text to be placed into title of the slide")]
            public TextStructure TitleText { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = false, Tooltip = "text to be placed in your slide")]
            public TextStructure SlideText { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = false, Tooltip = "slide layout")]
            public TextStructure SlideLayout{ get; set; } = new TextStructure(string.Empty);

        }
        public PowerPointInsertSlideCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            PowerPointWrapper ppWrapper = PowerPointManager.CurrentPowerPoint;
            int slidePosition = arguments.SlidePosition.Value;
            string titleText = arguments.TitleText.Value;
            string slideText = arguments.SlideText.Value;
            string slideLayoutText = arguments.SlideLayout.Value;

            try
            {
                ppWrapper.InsertSlide(slidePosition,titleText,slideText, slideLayoutText);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to insert slide. Message: {ex.Message}", ex);
            }
        }
    }
}
