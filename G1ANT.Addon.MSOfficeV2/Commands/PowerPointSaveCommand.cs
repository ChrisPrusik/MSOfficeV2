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
    [Command(Name = "powerpoint.save", Tooltip = "This command saves current presentation.", NeedsDelay = true, IsUnderConstruction = false)]

    public class PowerPointSaveCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "insert your desired save path")]
            public TextStructure PresName { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = false, Tooltip ="insert your desired save path")]
            public TextStructure Path { get; set; } = new TextStructure(string.Empty);

        }
        public PowerPointSaveCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            PowerPointWrapper ppWrapper = PowerPointManager.CurrentPowerPoint;     
            string path= arguments.Path.Value;
            string presName = arguments.PresName.Value;

            try
            { 
                ppWrapper.Save(presName,path);
            }
            catch (Exception ex)
            {

                throw new ApplicationException($"Error occured while trying to insert photo. Message: {ex.Message}", ex);

            }
        }
    }
}
