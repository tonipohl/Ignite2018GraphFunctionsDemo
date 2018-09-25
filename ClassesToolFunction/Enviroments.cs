using System;
using System.Collections.Generic;
using System.Text;

namespace ClassesToolFunction
{
    public class Enviroments
    {
        public static string GetEnvironmentVariable(string name)
        {
            return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
    }
}
