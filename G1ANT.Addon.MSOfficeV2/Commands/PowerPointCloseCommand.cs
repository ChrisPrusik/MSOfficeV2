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
    [Command(Name = "powerpoint.close", Tooltip = "This command closes current presentation.", NeedsDelay = true, IsUnderConstruction = false)]

    public class PowerPointCloseCommand : Command
    {
        public class Arguments : CommandArguments
        {
        }
        public PowerPointCloseCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            PowerPointWrapper ppWrapper = PowerPointManager.CurrentPowerPoint;     

            try
            { 
                ppWrapper.Close();
            }
            catch (Exception ex)
            {

                throw new ApplicationException($"Error occured while trying to insert photo. Message: {ex.Message}", ex);

            }
        }
    }
}
