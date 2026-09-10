# FAST CLASS

Workspace inicial do FAST CLASS, um copiloto de planejamento de aulas com IA para professores da educação básica brasileira.

## O que foi analisado

Os documentos fornecidos foram tratados como referências de produto, implementação e design:

- `Proposta de MVP - Gerador de Aulas.docx`: escopo do MVP, fluxo de criação e hipóteses do produto.
- `Guia de Desenvolvimento — AI Lesson Generator.docx`: arquitetura sugerida, integração com Supabase/OpenAI e requisitos de segurança.
- `Guia de Design System — Gerador de Aulas.docx`: tokens visuais, componentes, estados, responsividade e acessibilidade.

As decisões de implementação atuais priorizam o MVP em PT-BR e mantêm `workspace_id` como parte da arquitetura futura. Exemplos em inglês presentes nos documentos foram tratados como exemplos de referência, não como requisito de idioma do produto.

## Estado atual

O primeiro slice navegável já está configurado:

- shell responsivo com sidebar e navegação;
- dashboard em PT-BR;
- fluxo de criação de aula com modo Fast/Criativo;
- campos de tema, turma, duração, objetivos, metodologia e material;
- estado visual de geração e resultado estruturado;
- tokens do Design System aplicados em Tailwind/CSS;
- Supabase SSR preparado com `@supabase/ssr`, clientes browser/server e `proxy.ts`;
- placeholder seguro para `OPENAI_API_KEY`, sem chamadas externas ainda.

A geração exibida no protótipo é simulada. A integração real com OpenAI, persistência no Supabase, upload privado de arquivos, RLS e exportação PDF são os próximos incrementos.

## Stack

- Next.js 16 + App Router + TypeScript
- React 19
- Tailwind CSS
- Lucide React
- Supabase SSR/Auth/Postgres/Storage (base preparada)

## Rodando localmente

```bash
pnpm install
cp .env.example .env.local
pnpm dev
```

Validações disponíveis:

```bash
pnpm lint
pnpm exec tsc --noEmit --incremental false
pnpm build
```

Abra [http://localhost:3000](http://localhost:3000).

## Variáveis de ambiente

Preencha `.env.local` com as credenciais do projeto Supabase:

```env
NEXT_PUBLIC_SUPABASE_URL=
NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY=
OPENAI_API_KEY=
```

`OPENAI_API_KEY` é somente servidor e nunca deve receber o prefixo `NEXT_PUBLIC_` nem ser commitada.

## Próximos passos técnicos

1. Criar/conectar o projeto Supabase e confirmar o ambiente de desenvolvimento.
2. Modelar o primeiro schema multi-tenant (`workspaces`, membros, aulas e arquivos) com RLS.
3. Implementar a rota server-side de geração estruturada e validação do JSON.
4. Persistir rascunhos/histórico e adicionar upload privado com URLs assinadas.
5. Implementar exportação PDF e testes do fluxo crítico.

Não há migração SQL criada nesta etapa porque ainda faltam o projeto Supabase e as decisões finais de tenancy/roles.
