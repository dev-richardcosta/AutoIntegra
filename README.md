# Gerador Inteligente de Documentos (Word VBA)

Projeto VBA independente para Word que preenche todos os modelos em `Modelos` usando marcadores no formato `{{CAMPO}}`.

## Instalação

1. No Word, abra o VBA com `ALT+F11`.
2. Importe os `.bas` de `src` e `forms/frmPrincipal.frm`.
3. Salve o arquivo hospedeiro como `.docm` e execute `GerarDocumentos`.

Os modelos originais permanecem intactos; cópias de trabalho são processadas em `Temp` e as saídas ficam em `Gerados`.
