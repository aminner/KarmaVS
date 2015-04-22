// Guids.cs
// MUST match guids.h

using System;

namespace devcoach.Tools
{
    static class GuidList
    {
        public const string guidKarmaVsPkgString = "7ca5e40e-4946-4da8-be1f-b3bab3e8adfc";
        public const string guidKarmaVsUnitCmdSetString = "a91817b7-83bc-4d3b-bbac-67c87be2b5b5";
        public const string guidKarmaEnableDisableCmdSetString = "A9309E03-EDE4-4241-983E-413B3FFADC2D";
        public const string guidKarmaVsOptionsCmdSet = "875D357A-35F4-4142-B869-F7D664970086";


        public static readonly Guid guidKarmaVsPkg = new Guid(guidKarmaVsPkgString);
        public static readonly Guid guguidKarmaVsOptionsCmdSet = new Guid(guidKarmaVsOptionsCmdSet);
        public static readonly Guid guidKarmaVsUnitCmdSet = new Guid(guidKarmaVsUnitCmdSetString);
        public static readonly Guid guidKarmaEnableDisableCmdString = new Guid(guidKarmaEnableDisableCmdSetString);
    };
}