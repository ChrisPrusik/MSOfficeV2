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
    [Command(Name = "powerpoint.insertphoto", Tooltip = "This command inserts photo into current slide.", NeedsDelay = true, IsUnderConstruction = false)]

    public class PowerPointInsertPhotoCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip ="insert your photo's path")]
            public TextStructure PathToPhoto { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = true, Tooltip = "position of Photo relative to left")]
            public IntegerStructure Left { get; set; } = new IntegerStructure(1);

            [Argument(Required = true, Tooltip = "position of Photo relative to top")]
            public IntegerStructure Top { get; set; } = new IntegerStructure(1);

            [Argument(Required = true, Tooltip = "height of the photo")]
            public IntegerStructure Height { get; set; } = new IntegerStructure(1);

            [Argument(Required = true, Tooltip = "width of the photo")]
            public IntegerStructure Width { get; set; } = new IntegerStructure(1);

        }
        public PowerPointInsertPhotoCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            PowerPointWrapper ppWrapper = PowerPointManager.CurrentPowerPoint;
            string pathToPhoto= arguments.PathToPhoto.Value;
            int left = arguments.Left.Value;
            int top = arguments.Top.Value;
            int height = arguments.Height.Value;
            int width = arguments.Width.Value;

            try
            { 
                ppWrapper.InsertPhoto(pathToPhoto, left,top,height,width);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to insert photo. Message: {ex.Message}", ex);
            }
        }
    }
}
