1) cd SessionCOM   (current directory)
2) gacutil /i SessionUtility.dll
3) regasm.exe SessionUtility.dll /tlb:SessionUtility.tlb
4) regsvr32 SessionManager.dll
