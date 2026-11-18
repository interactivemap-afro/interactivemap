# Mapa CEsA — Atualização v2 (Portugal + Espanha)

## O que mudou
- **Seletor de projeto** (topo): *Projetos portugueses* / *Projetos ibéricos*.
  - Portugal → só os registos originais, vista inicial igual à anterior (Lisboa).
  - Ibérico → registos de Portugal + Espanha, vista da Península Ibérica.
- **Seletor de idioma** PT / ES / EN (inglês acrescentado na v3): traduz a interface, as categorias, as áreas de atuação e o texto "Sobre" quando existe tradução.
- Filtros, pesquisa e lista digital respeitam o projeto escolhido.
- Novas formas jurídicas: `cooperativa`, `empresa` e `medio` (meio de comunicação). Os ícones são provisórios, copiados de *cooperação* e *marca*.
- Etiqueta **Evento** para os registos com `"evento": true`.

## Compatibilidade com o embed
O endereço atual continua a funcionar **sem alterações**: sem parâmetros, o mapa abre como antes (PT, Portugal, Lisboa).
Parâmetros opcionais, por exemplo para o site do IMF-CSIC:

```html
<iframe src="https://SEU-URL/index.html?lang=es&projeto=iberico" ...></iframe>
```

| Parâmetro | Valores | Por defeito |
|---|---|---|
| `lang` | `pt`, `es`, `en` | `pt` |
| `projeto` | `pt`, `iberico` | `pt` |

## Dados (`data.json`)
- Os 86 registos originais **não foram alterados** (sem o campo `projeto`, contam como `pt`).
- Os novos registos têm `"projeto": "es"` e IDs a partir de **1001**.
- Campos novos (opcionais): `sobre_i18n` (`{"es": ..., "pt": ..., "en": ...}`), `formaJuridicaOriginal`, `areasOriginal`.

## Importar mais fichas DOCX
1. Põe os `.docx` em `fichas_es/` (formato: `NOMBRE:`, `Área actuación:`, `Forma jurídica:`, `Resumen:`, `Dirección:`, `Teléfono:`, `Email:`, `Website:`; opcionalmente `Resumen portugués:` e `Resumen inglés:`).
2. Acrescenta uma entrada em `tools/fichas_meta_es.json` (a chave é o nome do ficheiro .docx, sem extensão):
   `{"coords": [lat, lon], "logo": "LogoX.jpg", "evento": true}`. Tudo é opcional.
   Para as coordenadas: no Google Maps, botão direito na morada → copiar coordenadas.
3. Corre:
   ```bash
   pip install python-docx
   python3 tools/docx_to_json.py --data data.json --meta tools/fichas_meta_es.json fichas_es/*.docx
   ```
   Podes correr de novo sempre que quiseres: os registos com o mesmo nome são atualizados, não duplicados. O logótipo é extraído para `img/es/`.
4. Os avisos `⚠ sem coordenadas` indicam registos que ficam no JSON mas ainda não aparecem no mapa.

## Publicar no GitHub
1. Faz commit de: `index.html`, `data.json`, `img/es/`, `icons/` e `icons_filter/` (novos `cooperativa.svg` e `empresa.svg`), `tools/` e este ficheiro.
   - Por privacidade, `fichas_es/` pode ficar fora do repositório; acrescenta-a ao `.gitignore`, se preferires.
2. Depois do deploy, testa o URL sem parâmetros (tem de ficar igual ao anterior) e com `?lang=es&projeto=iberico`.

## v3 (15/11/2026) — correções da Yolanda + inglês
- **Africaespromociona removida** (web sequestrada, situação denunciada). Fica `"ignorar": true` em `fichas_meta_es.json`: se a ficha voltar a ser importada, o conversor apaga o registo em vez de o recriar. Apagar também `img/es/africaespromociona.png`, `fichas_es/Africaespromociona.docx` e `fichas_es/Logo Africaespromociona.png`.
- **Afrolab Basement**: telefone 677 517 889.
- **Otros Sures**: emails quinnymartinez168@gmail.com (principal) e santamonica@gencat.cat. O painel mostra agora todos os emails do campo `eMails`.
- **SOS Racisme**: morada confirmada pela entidade (rua renomeada; espaço de co-working). Resumo em castelhano.
- **Inglês (EN)**: interface, formas jurídicas e áreas de atuação. Os resumos das 24 entidades de Espanha estão em ES, PT e EN (`sobre_i18n`).
- **Correções manuais em `fichas_meta_es.json`** (`telefone`, `eMail`, `eMails`, `sobre_i18n`, `notaMorada`) têm prioridade sobre a ficha DOCX. Ao reimportar, as traduções PT/EN já existentes mantêm-se.
- Embed em inglês: `index.html?lang=en&projeto=iberico`.
