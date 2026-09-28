# Dump members of selected interfaces from CorelDraw.tlb (offline, no CorelDRAW needed).
# usage: powershell -ExecutionPolicy Bypass -File tools\tlb_dump.ps1 Curve Node Segment SubPath Shape ShapeRange Page Document Layer
param(
  [string[]]$Want = @("Curve","Node","Segment","SubPath"),
  [string]$Tlb = "C:\Program Files (x86)\CorelDRAW X4\Programs\CorelDraw.tlb",
  [string]$Guid = "",
  [int]$Major = 14,
  [int]$Minor = 0
)

$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using CT = System.Runtime.InteropServices.ComTypes;

public class TlbDump {
  [DllImport("oleaut32.dll", CharSet=CharSet.Unicode, PreserveSig=false)]
  static extern void LoadTypeLib(string fileName, out CT.ITypeLib tlb);

  [DllImport("oleaut32.dll", CharSet=CharSet.Unicode, PreserveSig=false)]
  static extern void LoadRegTypeLib(ref Guid rguid, short major, short minor, int lcid, out CT.ITypeLib tlb);

  public static void DumpReg(string guid, short major, short minor, string[] want) {
    Guid g = new Guid(guid);
    CT.ITypeLib tlb;
    LoadRegTypeLib(ref g, major, minor, 0, out tlb);
    Walk(tlb, want);
  }

  public static void Dump(string path, string[] want) {
    CT.ITypeLib tlb;
    LoadTypeLib(path, out tlb);
    Walk(tlb, want);
  }

  static void Walk(CT.ITypeLib tlb, string[] want) {
    int n = tlb.GetTypeInfoCount();
    for (int i = 0; i < n; i++) {
      string nm, doc, help;
      int ctx;
      tlb.GetDocumentation(i, out nm, out doc, out ctx, out help);
      if (nm == null) continue;
      bool match = false;
      for (int w = 0; w < want.Length; w++)
        if (want[w] == "*") match = true;
        else if (string.Equals(nm, want[w], StringComparison.OrdinalIgnoreCase)) match = true;
      if (!match) continue;

      CT.ITypeInfo ti;
      tlb.GetTypeInfo(i, out ti);
      IntPtr pta;
      ti.GetTypeAttr(out pta);
      CT.TYPEATTR ta = (CT.TYPEATTR)Marshal.PtrToStructure(pta, typeof(CT.TYPEATTR));
      Console.WriteLine("=== " + nm + "  (" + ta.typekind + ", funcs=" + ta.cFuncs + ") ===");
      for (int f = 0; f < ta.cFuncs; f++) {
        IntPtr pfd;
        ti.GetFuncDesc(f, out pfd);
        CT.FUNCDESC fd = (CT.FUNCDESC)Marshal.PtrToStructure(pfd, typeof(CT.FUNCDESC));
        string[] names = new string[fd.cParams + 1];
        int got;
        ti.GetNames(fd.memid, names, names.Length, out got);
        string line = "   " + names[0];
        if (fd.invkind == CT.INVOKEKIND.INVOKE_PROPERTYGET) line += "   [get]";
        else if (fd.invkind == CT.INVOKEKIND.INVOKE_PROPERTYPUT) line += "   [put]";
        else line += "   (params=" + fd.cParams + ", opt=" + fd.cParamsOpt + ")";
        Console.WriteLine(line);
        ti.ReleaseFuncDesc(pfd);
      }
      ti.ReleaseTypeAttr(pta);
    }
  }
}
"@

if ($Guid -ne "") {
  [TlbDump]::DumpReg($Guid, [int16]$Major, [int16]$Minor, $Want)
} else {
  [TlbDump]::Dump($Tlb, $Want)
}