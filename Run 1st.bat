copy bin\sys32\hookmenu.ocx %systemroot%\system32\hookMenu.ocx
copy bin\sys32\SubclassingSink.tlb %systemroot%\system32\SubclassingSink.tlb
regsvr32 /s %systemroot%\system32\SubclassingSink.tlb