Exemplo de Preview Personalizado para FastReport
-

Criei este exemplo de preview personalizado para Fast Report, há alguns anos.  
Só adicionar ao seu projeto e ajustar o layout  

como usar  

    fRel_Preview := TfRel_Preview.Create(nil);  
    fRel_Preview.FastRep.Clear;  
    fRel_Preview.FastRep.LoadFromFile(ArquivoFR3);  
    fRel_Preview.FastRep.PrepareReport;  
    fRel_Preview.FastRep.ShowReport(False);  
    fRel_Preview.ShowModal;  
    FreeAndNil(fRel_Preview);  
