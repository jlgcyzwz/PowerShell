﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write(InputSender.Send(InputSender.MouseWheelDirection.Vertical, 2));
        }
    }
}
