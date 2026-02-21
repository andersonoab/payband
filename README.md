# Compensation Bands  
Uso interno | Ferramenta interativa de análise de Compa Ratio  

---

## 1. Visão Geral

O **Compensation Bands** é uma aplicação web estática (HTML + CSS + JS) para análise salarial baseada em:

- Comparação com a aba **Pay Bands** (quando existente no Excel)
- Cálculo estimado automático (fallback) quando não houver tabela oficial
- Compa Ratio (Salário ÷ P100)
- Classificação automática:
  - Abaixo (<80)
  - Dentro (80-120)
  - Acima (>120)

A ferramenta foi construída para apoiar decisões de People Analytics, revisões salariais, calibração e auditorias internas.

Versão atual: **v4.2**

---

## 2. Principais Funcionalidades

### 2.1 Importação Inteligente de Excel

- Detecta automaticamente a aba principal (ex: Excel Output)
- Se existir aba **Pay Bands**, utiliza:
  - Pay Positioning (80 / 100 / 120)
  - Sonova Level (A-J)
  - Currency
  - Comparação por:
    - Position Role Family (externalName)
- Caso não exista Pay Bands:
  - Estima P80, P100, P120 por grupo + level + moeda

---

### 2.2 Filtros Dinâmicos

Filtros fixos:

- Busca textual
- Job Family
- Pay Band
- Sonova Level
- Status
- Currency
- Compa mínimo / máximo

Filtros dinâmicos:

- Qualquer nova coluna adicionada na aba Excel Output:
  - Aparece automaticamente em "Colunas extras"
  - Se marcada, passa a:
    - Aparecer na tabela
    - Virar filtro automaticamente

Regras de filtro:

- Até 30 valores únicos → Select (igualdade)
- Muitos valores → Campo texto (contém)
- Cada filtro possui botão **X** para limpeza rápida

---

### 2.3 Ordenação Inteligente

Clique no cabeçalho da coluna:

- Primeiro clique → Crescente (ASC)
- Segundo clique → Decrescente (DESC)

Funciona para:

- Texto
- Valores monetários
- Compa
- Colunas extras
- P80 / P100 / P120
- Status

Indicador visual ASC / DESC no cabeçalho.

---

### 2.4 Exportações

Exportar TXT  
Exporta apenas os registros filtrados.

Exportar Excel  
- Exporta apenas os registros filtrados
- Inclui todas as colunas
- Mantém:
  - BaseSalary
  - P80 / P100 / P120
  - Compa
  - Fonte da faixa
  - Todas as colunas extras detectadas

---

## 3. Estrutura do Projeto
index.html
style.css
app.js

Tecnologia:

- HTML puro
- CSS customizado
- JavaScript Vanilla
- Biblioteca XLSX (SheetJS via CDN)

Sem backend.  
Sem dependências de servidor.  
Funciona localmente.

---

## 4. Lógica de Cálculo

### 4.1 Prioridade da Faixa

1. Se existir aba "Pay Bands":
   - Usa valores oficiais 80/100/120
2. Se não existir:
   - Calcula percentis internos por grupo
   - Fallback 0.8 / 1.2 quando grupo pequeno

### 4.2 Compa Ratio

Classificação:

- < P80 → Abaixo
- Entre P80 e P120 → Dentro
- > P120 → Acima

---

## 5. Colunas Extras (Arquitetura Dinâmica)

O sistema detecta automaticamente todas as colunas que não fazem parte do núcleo:

Colunas núcleo:

- EmployeeName
- EmployeeId
- JobFamily
- PayBand
- Level
- Currency
- BaseSalary

Qualquer outra coluna:

- Vira "Coluna extra"
- Pode ser exibida
- Pode virar filtro
- Pode ser exportada

Exemplos típicos:

- CDC
- Chefia
- Tipo
- Location
- Cost Center
- Functional Area

---

## 6. Experiência do Usuário

Melhorias implementadas na v4.2:

- Botão X para limpar filtros individuais
- Ordenação clicável
- Filtros dinâmicos por coluna marcada
- Exportação Excel conforme filtros aplicados
- Persistência via localStorage

---

## 7. Uso

1. Abrir index.html
2. Clicar em "Importar Excel"
3. Selecionar arquivo .xlsx
4. Aplicar filtros
5. Ordenar colunas conforme necessidade
6. Exportar resultado final

---

## 8. Governança e Uso Interno

Ferramenta desenvolvida para:

- People Analytics
- Revisão salarial
- Auditorias internas
- Calibração de bandas
- Simulações estratégicas

Uso interno Sonova.

---

## 9. Roadmap Futuro (Opcional)

Possíveis evoluções:

- Exportação PowerPoint executivo
- Gráficos automáticos por faixa
- Heatmap por Job Family
- Simulador de ajuste salarial
- Integração com dashboard BI

---

## 10. Versão Atual

v4.2  
Base oficial para evoluções futuras.

---
Anderson Marinho
