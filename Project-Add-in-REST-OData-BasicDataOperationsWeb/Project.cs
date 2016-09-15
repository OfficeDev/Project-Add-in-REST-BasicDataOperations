// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Project_Add_in_REST_BasicDataOp
{
    //Helper calss to represent a project to  bind to the grid view.
    public class Project
    {
        public String id { get; set; }
        public String name { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public bool isCheckedOut { get; set; }
    }
}
