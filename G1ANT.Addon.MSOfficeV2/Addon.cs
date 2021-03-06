/**
* Copyright(C) G1ANT Ltd, All rights reserved
* Solution G1ANT.Addon.MSOffice, Project G1ANT.Addon.MSOffice
* www.g1ant.com
*
* Licensed under the G1ANT license.
* See License.txt file in the project root for full license information.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;
using System.Drawing;

namespace G1ANT.Addon.MSOfficeV2
{
    [Addon(Name = "msofficev2",
        Tooltip = "MSOfficeV2 Commands")]
    [Copyright(Author = "G1ANT LTD", Copyright = "G1ANT LTD", Email = "hi@g1ant.com", Website = "www.g1ant.com")]
    [License(Type = "LGPL", ResourceName = "License.txt")]
    [CommandGroup(Name = "excel", Tooltip = "Command connected with creating, editing and generally working on excel")]
    [CommandGroup(Name = "word",  Tooltip = "Command connected with creating, editing and generally working on word")]
    [CommandGroup(Name = "powerpoint", Tooltip = "Command connected with creating, editing and generally working on powerpoint")]
    [CommandGroup(Name = "outlook", Tooltip = "Command connected with creating, editing and generally working on outlook")]

    public class Addon : Language.Addon
    {
        
    }
}
