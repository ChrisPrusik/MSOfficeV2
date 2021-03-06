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
    [Command(Name = "outlook.send",Tooltip = "This command to send a prepared earlier message by 'outlook.newmessage'.")]
    public class OutlookSendCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            [Argument]
            public OutlookMailStructure Mail { get; set; }
        }
        public OutlookSendCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager != null)
            {
                outlookManager.Send(arguments.Mail?.Value);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.BooleanStructure(true));
            }
            else
                throw new NullReferenceException("Current Outlook is not set.");
        }
    }
}
