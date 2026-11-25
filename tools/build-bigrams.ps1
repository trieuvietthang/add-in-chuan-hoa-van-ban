#requires -Version 5.1
# Dung bang can am tiet (bigram) tu ngu lieu cau DUNG. Chay thu cong khi muon cap nhat du lieu; ket qua nam trong config/ (khong chay trong add-in).
# Nguon: bmd1905/error-correction-vi (Apache-2.0), cot correct_text. Chi giu cac cap xuat hien >= MinCount lan.
param(
    [Parameter(Mandatory)][string]$Csv,
    [string]$Out = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config\bigram.vi.txt'),
    [int]$MinCount = 5
)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System; using System.IO; using System.Text; using System.Collections.Generic; using Microsoft.VisualBasic.FileIO;
public static class BG {
  public static void Run(string csv, string outp, int minc) {
    var uni = new Dictionary<string,int>(); var bi = new Dictionary<string,int>();
    using (var p = new TextFieldParser(csv, Encoding.UTF8)) {
      p.TextFieldType = FieldType.Delimited; p.SetDelimiters(","); p.HasFieldsEnclosedInQuotes = true;
      p.ReadFields();
      while (!p.EndOfData) {
        string[] f; try { f = p.ReadFields(); } catch { continue; }
        if (f.Length < 2) continue;
        string s = f[1].Normalize(NormalizationForm.FormC).ToLowerInvariant();
        string prev = null; var sb = new StringBuilder();
        for (int i = 0; i <= s.Length; i++) {
          char c = i < s.Length ? s[i] : ' ';
          if (char.IsLetter(c)) { sb.Append(c); continue; }
          if (sb.Length > 0) {
            string w = sb.ToString(); sb.Clear();
            int n; uni.TryGetValue(w, out n); uni[w] = n + 1;
            if (prev != null) { string k = prev + " " + w; bi.TryGetValue(k, out n); bi[k] = n + 1; }
            prev = w;
          }
          // chi dau cau/so lam dut cap; khoang trang giu nguyen mach
          if (c != ' ' && (int)c != 160) prev = null;
        }
      }
    }
    using (var o = new StreamWriter(outp, false, new UTF8Encoding(false))) {
      o.WriteLine("; Bang tan suat am tiet / cap am tiet tu ngu lieu cau dung (Apache-2.0, bmd1905/error-correction-vi). Sinh boi tools/build-bigrams.ps1.");
      o.WriteLine("; Dong 'u<TAB>am_tiet<TAB>so_lan' va 'b<TAB>am_tiet1 am_tiet2<TAB>so_lan'.");
      foreach (var kv in uni) if (kv.Value >= minc) o.WriteLine("u\t" + kv.Key + "\t" + kv.Value);
      foreach (var kv in bi) if (kv.Value >= minc) o.WriteLine("b\t" + kv.Key + "\t" + kv.Value);
    }
  }
}
'@ -ReferencedAssemblies 'Microsoft.VisualBasic'
[BG]::Run([string]$Csv, [string]$Out, $MinCount)
"{0} dong, {1:N1} MB" -f (Get-Content $Out -Encoding UTF8 | Measure-Object -Line).Lines, ((Get-Item $Out).Length / 1MB)
