//------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft">
//      Copyright (c) Microsoft. All Rights Reserved.
//      Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.
// </copyright>
// <summary>
// MIM Configuration Documenter Main Program
// </summary>
//------------------------------------------------------------------------------------------------------------------------------------------

namespace MIMConfigDocumenter
{
    using System;
    using System.Globalization;
    using System.Reflection;

    /// <summary>
    /// MIM Configuration Documenter Entry Point
    /// </summary>
    public class Program
    {
        /// <summary>
        /// MIM Configuration Documenter Entry Point.
        /// </summary>
        /// <param name="args">The command-line arguments.</param>
        public static void Main(string[] args)
        {
            if (args == null || args.Length < 2)
            {
                string errorMsg = string.Format(CultureInfo.CurrentUICulture, "Usage: {0} {1} {2}.", new object[] { Assembly.GetExecutingAssembly().GetName().Name, "{Pilot / Target Config Folder}", "{Production / Reference / Baseline Config Folder}" });
                throw new ArgumentException(errorMsg, "args");
            }

            if (args.Length == 3)
            {
                switch (args[2])
                {
                    case "SyncOnly":
                        var syncDocumenter = new MIMSyncConfigDocumenter(args[0], args[1]);
                        syncDocumenter.GenerateReport();
                        return;

                    case "ServiceOnly":
                        var serviceDocumenter = new MIMServiceConfigDocumenter(args[0], args[1]);
                        serviceDocumenter.GenerateReport();
                        return;
                }
            }

            var documenter = new MIMConfigDocumenter(args[0], args[1]);
            documenter.GenerateReport();
        }
    }
}
