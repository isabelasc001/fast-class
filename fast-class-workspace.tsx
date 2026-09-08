"use client";

import {
  BookOpen,
  Check,
  ChevronRight,
  Clock3,
  FileText,
  FolderOpen,
  LayoutDashboard,
  LoaderCircle,
  Menu,
  Plus,
  Search,
  Settings2,
  Sparkles,
  Upload,
} from "lucide-react";
import { useState, type FormEvent } from "react";

type View = "dashboard" | "builder" | "lessons" | "result";
type LessonMode = "fast" | "criativo";

const navItems: Array<{ label: string; view: Exclude<View, "result" | "builder">; icon: typeof LayoutDashboard }> = [
  { label: "Visão geral", view: "dashboard", icon: LayoutDashboard },
  { label: "Minhas aulas", view: "lessons", icon: BookOpen },
];

const inputClass =
  "focus-ring mt-2 w-full rounded-[var(--radius-sm)] border border-fast-border bg-fast-surface px-3.5 py-3 text-sm text-fast-heading shadow-[var(--shadow-sm)] outline-none transition placeholder:text-fast-text-muted focus:border-fast-primary";

function Brand() {
  return (
    <div className="flex items-center gap-3">
      <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-fast-primary text-white shadow-[var(--shadow-sm)]">
        <Sparkles size={19} strokeWidth={2} aria-hidden="true" />
      </div>
      <div>
        <p className="font-display text-xl leading-none text-fast-heading">FAST CLASS</p>
        <p className="mt-1 text-[11px] font-medium uppercase tracking-[0.18em] text-fast-text-muted">
          workspace pedagógico
        </p>
      </div>
    </div>
  );
}

function Sidebar({ view, onNavigate }: { view: View; onNavigate: (view: View) => void }) {
  return (
    <aside className="hidden min-h-screen w-64 shrink-0 border-r border-fast-border-soft bg-fast-surface-muted px-5 py-6 lg:flex lg:flex-col">
      <Brand />
      <div className="mt-10 flex items-center gap-3 rounded-xl border border-fast-border-soft bg-fast-surface px-3 py-3 shadow-[var(--shadow-sm)]">
        <div className="flex h-9 w-9 items-center justify-center rounded-full bg-[#f1d6bf] text-sm font-semibold text-fast-primary">
          JM
        </div>
        <div className="min-w-0">
          <p className="truncate text-sm font-semibold text-fast-heading">Jaisson Monteiro</p>
          <p className="text-xs text-fast-text-muted">Workspace pessoal</p>
        </div>
      </div>
      <nav className="mt-9 space-y-1" aria-label="Navegação principal">
        {navItems.map(({ label, view: targetView, icon: Icon }) => {
          const active = view === targetView;
          return (
            <button
              key={targetView}
              type="button"
              onClick={() => onNavigate(targetView)}
              className={`focus-ring flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left text-sm font-medium transition ${
                active
                  ? "bg-[#ead7c9] text-fast-primary"
                  : "text-fast-text-muted hover:bg-fast-surface hover:text-fast-heading"
              }`}
              aria-current={active ? "page" : undefined}
            >
              <Icon size={18} strokeWidth={1.8} aria-hidden="true" />
              {label}
            </button>
          );
        })}
        <button
          type="button"
          onClick={() => onNavigate("builder")}
          className={`focus-ring flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left text-sm font-medium transition ${
            view === "builder"
              ? "bg-[#ead7c9] text-fast-primary"
              : "text-fast-text-muted hover:bg-fast-surface hover:text-fast-heading"
          }`}
        >
          <Sparkles size={18} strokeWidth={1.8} aria-hidden="true" />
          Criar com IA
        </button>
      </nav>
      <div className="mt-auto border-t border-fast-border-soft pt-4">
        <button type="button" className="focus-ring flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium text-fast-text-muted transition hover:bg-fast-surface hover:text-fast-heading">
          <Settings2 size={18} strokeWidth={1.8} aria-hidden="true" />
          Configurações
        </button>
      </div>
    </aside>
  );
}

function MobileHeader({ onMenu }: { onMenu: () => void }) {
  return (
    <header className="flex items-center justify-between border-b border-fast-border-soft bg-fast-surface px-5 py-4 lg:hidden">
      <Brand />
      <button type="button" onClick={onMenu} className="focus-ring rounded-lg p-2 text-fast-heading" aria-label="Abrir menu">
        <Menu size={22} strokeWidth={1.8} aria-hidden="true" />
      </button>
    </header>
  );
}

function Dashboard({ onCreate }: { onCreate: () => void }) {
  return (
    <div className="space-y-10">
      <section className="relative overflow-hidden rounded-[var(--radius-xl)] bg-[#f1e3d6] px-6 py-8 sm:px-10 sm:py-10">
        <div className="relative z-10 max-w-2xl">
          <p className="mb-4 text-sm font-semibold uppercase tracking-[0.16em] text-fast-primary">terça-feira, 8 de setembro</p>
          <h1 className="font-display max-w-xl text-4xl leading-[1.08] text-fast-heading sm:text-5xl">
            Vamos preparar uma aula que faça sentido para a sua turma.
          </h1>
          <p className="mt-5 max-w-lg text-base leading-7 text-fast-text">
            Conte o que você quer ensinar, compartilhe seus materiais e deixe o FAST CLASS organizar o próximo passo.
          </p>
          <button
            type="button"
            onClick={onCreate}
            className="focus-ring mt-7 inline-flex items-center gap-2 rounded-lg bg-fast-primary px-5 py-3 text-sm font-semibold text-white shadow-[var(--shadow-sm)] transition hover:bg-fast-primary-hover"
          >
            <Plus size={18} strokeWidth={2} aria-hidden="true" />
            Criar nova aula
          </button>
        </div>
        <div className="pointer-events-none absolute -right-10 -top-16 hidden h-72 w-72 rounded-full border-[26px] border-[#e2c4ad] opacity-70 sm:block" />
        <div className="pointer-events-none absolute bottom-[-70px] right-24 hidden h-48 w-48 rounded-full border-[22px] border-[#e9c19c] opacity-50 sm:block" />
        <Sparkles className="absolute bottom-8 right-10 text-fast-accent opacity-80 sm:right-16" size={42} strokeWidth={1.4} aria-hidden="true" />
      </section>

      <section className="space-y-5" aria-labelledby="recent-lessons-title">
        <div className="flex items-end justify-between gap-4">
          <div>
            <p className="text-sm font-medium text-fast-text-muted">Continuidade</p>
            <h2 id="recent-lessons-title" className="font-display mt-1 text-3xl text-fast-heading">Aulas recentes</h2>
          </div>
          <button type="button" className="focus-ring hidden items-center gap-1 text-sm font-semibold text-fast-primary sm:flex">
            Ver todas <ChevronRight size={16} aria-hidden="true" />
          </button>
        </div>
        <div className="grid gap-4 md:grid-cols-2">
          <article className="group rounded-[var(--radius-lg)] border border-fast-border-soft bg-fast-surface p-5 shadow-[var(--shadow-sm)] transition hover:border-fast-border hover:shadow-[var(--shadow-md)]">
            <div className="flex items-start justify-between gap-4">
              <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-[#f6e6d9] text-fast-primary">
                <BookOpen size={20} strokeWidth={1.8} aria-hidden="true" />
              </div>
              <span className="rounded-full bg-[#e7f0e8] px-2.5 py-1 text-xs font-semibold text-fast-success">rascunho</span>
            </div>
            <h3 className="mt-5 text-lg font-semibold text-fast-heading">Present Perfect: Travel Experiences</h3>
            <p className="mt-2 text-sm text-fast-text-muted">Inglês · 9º ano · 60 min</p>
            <div className="mt-6 flex items-center justify-between border-t border-fast-border-soft pt-4 text-xs text-fast-text-muted">
              <span className="inline-flex items-center gap-1.5"><Clock3 size={14} aria-hidden="true" /> editada há 2 horas</span>
              <button type="button" className="focus-ring font-semibold text-fast-primary">Abrir aula</button>
            </div>
          </article>
          <button type="button" onClick={onCreate} className="focus-ring flex min-h-56 flex-col items-center justify-center rounded-[var(--radius-lg)] border border-dashed border-fast-border bg-transparent px-6 text-center transition hover:border-fast-primary hover:bg-[#fffaf6]">
            <span className="flex h-11 w-11 items-center justify-center rounded-full bg-[#f6e6d9] text-fast-primary"><Plus size={20} aria-hidden="true" /></span>
            <span className="mt-4 text-sm font-semibold text-fast-heading">Começar uma nova aula</span>
            <span className="mt-1 text-sm text-fast-text-muted">O FAST CLASS ajuda a estruturar o pedido.</span>
          </button>
        </div>
      </section>
    </div>
  );
}

function Builder({ onBack, onGenerated }: { onBack: () => void; onGenerated: () => void }) {
  const [mode, setMode] = useState<LessonMode>("fast");
  const [generating, setGenerating] = useState(false);

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setGenerating(true);
    window.setTimeout(() => {
      setGenerating(false);
      onGenerated();
    }, 1100);
  }

  return (
    <div className="space-y-8">
      <div className="flex items-start justify-between gap-4">
        <div>
          <button type="button" onClick={onBack} className="focus-ring mb-4 inline-flex items-center gap-1.5 text-sm font-semibold text-fast-primary"><ChevronRight className="rotate-180" size={16} aria-hidden="true" /> Voltar</button>
          <p className="text-sm font-medium text-fast-text-muted">Nova aula · etapa 1 de 6</p>
          <h1 className="font-display mt-1 text-4xl text-fast-heading">Estruture o seu pedido</h1>
          <p className="mt-3 max-w-2xl text-base leading-7 text-fast-text">Algumas respostas ajudam a transformar uma boa ideia em uma aula possível para a sua turma.</p>
        </div>
        <span className="hidden rounded-full bg-[#f8e6d9] px-3 py-1.5 text-xs font-semibold text-fast-primary sm:inline-flex">Fast</span>
      </div>

      <div className="h-1.5 overflow-hidden rounded-full bg-[#eadfd5]" aria-label="Progresso da criação"><div className="h-full w-1/6 rounded-full bg-fast-primary" /></div>

      <div className="grid gap-6 xl:grid-cols-[minmax(0,1fr)_300px]">
        <form onSubmit={handleSubmit} className="rounded-[var(--radius-lg)] border border-fast-border-soft bg-fast-surface p-5 shadow-[var(--shadow-sm)] sm:p-8">
          <div className="grid gap-6 sm:grid-cols-2">
            <label className="sm:col-span-2 text-sm font-semibold text-fast-heading">Disciplina e tema
              <input className={inputClass} name="topic" placeholder="Ex.: Ciências — ciclo da água" required />
            </label>
            <label className="text-sm font-semibold text-fast-heading">Série ou nível
              <select className={inputClass} name="grade" defaultValue="" required>
                <option value="" disabled>Selecione a turma</option>
                <option>6º ano</option><option>7º ano</option><option>8º ano</option><option>9º ano</option>
              </select>
            </label>
            <label className="text-sm font-semibold text-fast-heading">Duração
              <select className={inputClass} name="duration" defaultValue="60 min">
                <option>30 min</option><option>45 min</option><option>60 min</option><option>90 min</option>
              </select>
            </label>
            <label className="sm:col-span-2 text-sm font-semibold text-fast-heading">O que os estudantes devem conseguir fazer?
              <textarea className={`${inputClass} min-h-28 resize-y`} name="goal" placeholder="Descreva um resultado observável ao final da aula." required />
            </label>
            <label className="text-sm font-semibold text-fast-heading">Metodologia
              <select className={inputClass} name="methodology" defaultValue="equilibrada">
                <option value="equilibrada">Equilibrada</option><option value="projetos">Aprendizagem por projetos</option><option value="colaborativa">Colaborativa</option><option value="investigativa">Investigativa</option>
              </select>
            </label>
            <label className="text-sm font-semibold text-fast-heading">Quantidade de alunos
              <input className={inputClass} name="students" type="number" min="1" placeholder="Ex.: 28" />
            </label>
            <div className="sm:col-span-2">
              <p className="text-sm font-semibold text-fast-heading">Material de apoio <span className="font-normal text-fast-text-muted">(opcional)</span></p>
              <label className="focus-ring mt-2 flex cursor-pointer flex-col items-center justify-center rounded-[var(--radius-md)] border border-dashed border-fast-border bg-[#fffaf6] px-5 py-7 text-center transition hover:border-fast-primary">
                <Upload className="text-fast-primary" size={22} strokeWidth={1.8} aria-hidden="true" />
                <span className="mt-3 text-sm font-semibold text-fast-heading">Arraste um arquivo ou clique para adicionar</span>
                <span className="mt-1 text-xs text-fast-text-muted">PDF, DOCX, imagem ou texto · até 10 MB</span>
                <input className="sr-only" type="file" accept=".pdf,.docx,.txt,.png,.jpg,.jpeg" />
              </label>
            </div>
          </div>

          <div className="mt-8 flex flex-col-reverse justify-between gap-3 border-t border-fast-border-soft pt-5 sm:flex-row sm:items-center">
            <p className="text-xs text-fast-text-muted">Você poderá revisar tudo antes de usar.</p>
            <button disabled={generating} type="submit" className="focus-ring inline-flex items-center justify-center gap-2 rounded-lg bg-fast-primary px-5 py-3 text-sm font-semibold text-white shadow-[var(--shadow-sm)] transition hover:bg-fast-primary-hover disabled:cursor-wait disabled:opacity-80">
              {generating ? <LoaderCircle className="animate-spin" size={17} aria-hidden="true" /> : <Sparkles size={17} aria-hidden="true" />}
              {generating ? "Criando sua aula..." : "Gerar aula"}
            </button>
          </div>
        </form>

        <aside className="space-y-4">
          <div className="rounded-[var(--radius-lg)] border border-fast-border-soft bg-fast-surface p-5 shadow-[var(--shadow-sm)]">
            <div className="flex items-center justify-between">
              <div><p className="text-sm font-semibold text-fast-heading">Modo de criação</p><p className="mt-1 text-xs leading-5 text-fast-text-muted">Você pode mudar depois.</p></div>
              <Sparkles size={18} className="text-fast-accent" aria-hidden="true" />
            </div>
            <div className="mt-5 space-y-2">
              {(["fast", "criativo"] as LessonMode[]).map((option) => {
                const active = mode === option;
                return <button key={option} type="button" onClick={() => setMode(option)} className={`focus-ring w-full rounded-lg border px-3.5 py-3 text-left transition ${active ? "border-fast-primary bg-[#fff7f1]" : "border-fast-border-soft hover:border-fast-border"}`} aria-pressed={active}>
                  <span className="flex items-center justify-between"><span className="text-sm font-semibold capitalize text-fast-heading">{option}</span>{active && <Check size={16} className="text-fast-primary" aria-hidden="true" />}</span>
                  <span className="mt-1 block text-xs leading-5 text-fast-text-muted">{option === "fast" ? "Poucas entradas e uma saída utilizável rapidamente." : "Mais contexto, alternativas e refinamento em etapas."}</span>
                </button>;
              })}
            </div>
          </div>
          <div className="rounded-[var(--radius-lg)] bg-fast-surface-muted p-5">
            <p className="text-xs font-semibold uppercase tracking-[0.14em] text-fast-primary">O que vem depois</p>
            <ol className="mt-4 space-y-3 text-sm text-fast-text">
              {["Plano de aula e objetivos", "Atividade e materiais", "Avaliação e gabarito", "Sugestão de recuperação"].map((step, index) => <li key={step} className="flex gap-3"><span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-[#e9d4c4] text-[11px] font-bold text-fast-primary">{index + 1}</span><span>{step}</span></li>)}
            </ol>
          </div>
        </aside>
      </div>
    </div>
  );
}

function LessonResult({ onBack }: { onBack: () => void }) {
  return (
    <div className="space-y-8">
      <div className="flex flex-col justify-between gap-5 sm:flex-row sm:items-start">
        <div><button type="button" onClick={onBack} className="focus-ring mb-4 inline-flex items-center gap-1.5 text-sm font-semibold text-fast-primary"><ChevronRight className="rotate-180" size={16} aria-hidden="true" /> Voltar ao dashboard</button><p className="text-sm font-medium text-fast-success">✓ Aula criada · pronta para revisão</p><h1 className="font-display mt-1 text-4xl text-fast-heading">Present Perfect: Travel Experiences</h1><p className="mt-2 text-sm text-fast-text-muted">Inglês · 9º ano · 60 minutos · Fast</p></div>
        <div className="flex gap-2"><button type="button" className="focus-ring rounded-lg border border-fast-border bg-fast-surface px-4 py-2.5 text-sm font-semibold text-fast-heading transition hover:bg-fast-surface-muted">Salvar</button><button type="button" className="focus-ring inline-flex items-center gap-2 rounded-lg bg-fast-primary px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-fast-primary-hover"><FileText size={16} aria-hidden="true" /> Exportar PDF</button></div>
      </div>
      <div className="grid gap-6 xl:grid-cols-[minmax(0,1fr)_300px]">
        <article className="space-y-6 rounded-[var(--radius-lg)] border border-fast-border-soft bg-fast-surface p-6 shadow-[var(--shadow-sm)] sm:p-8">
          <section><p className="text-xs font-semibold uppercase tracking-[0.14em] text-fast-primary">Objetivos de aprendizagem</p><h2 className="mt-2 text-xl font-semibold text-fast-heading">Ao final da aula, os estudantes serão capazes de:</h2><ul className="mt-3 list-disc space-y-2 pl-5 text-sm leading-6 text-fast-text"><li>Relatar experiências de viagem usando o present perfect.</li><li>Reconhecer a diferença entre experiências e acontecimentos concluídos.</li></ul></section>
          <section className="border-t border-fast-border-soft pt-6"><p className="text-xs font-semibold uppercase tracking-[0.14em] text-fast-primary">Sequência didática</p><div className="mt-4 space-y-5"><LessonStep number="01" title="Warm-up" duration="10 min" text="Converse com a turma sobre lugares que já visitaram e registre expressões recorrentes." /><LessonStep number="02" title="Vocabulary in context" duration="15 min" text="Apresente cartões com destinos e experiências; os alunos associam imagem, verbo e frase." /><LessonStep number="03" title="Speaking activity" duration="20 min" text="Em duplas, os estudantes trocam perguntas usando Have you ever...? e registram respostas." /></div></section>
          <section className="border-t border-fast-border-soft pt-6"><p className="text-xs font-semibold uppercase tracking-[0.14em] text-fast-primary">Avaliação rápida</p><p className="mt-2 text-sm leading-6 text-fast-text">Peça que cada estudante escreva três frases sobre experiências próprias e uma frase sobre algo que ainda não fez.</p></section>
        </article>
        <aside className="space-y-4"><div className="rounded-[var(--radius-lg)] border border-fast-border-soft bg-fast-surface p-5 shadow-[var(--shadow-sm)]"><div className="flex items-center gap-2 text-sm font-semibold text-fast-heading"><Sparkles size={17} className="text-fast-accent" aria-hidden="true" /> Assistente de IA</div><p className="mt-3 text-sm leading-6 text-fast-text-muted">O que você quer ajustar nesta aula?</p><div className="mt-4 space-y-2">{["Simplificar instruções", "Criar exercício de fala", "Adaptar para iniciantes", "Adicionar dever de casa"].map((action) => <button key={action} type="button" className="focus-ring w-full rounded-lg border border-fast-border-soft px-3 py-2.5 text-left text-xs font-semibold text-fast-text transition hover:border-fast-primary hover:text-fast-primary">{action}</button>)}</div><button type="button" className="focus-ring mt-4 flex w-full items-center gap-2 border-t border-fast-border-soft pt-4 text-left text-sm text-fast-text-muted"><Search size={15} aria-hidden="true" /> Perguntar outra coisa...</button></div><div className="rounded-[var(--radius-lg)] bg-fast-surface-muted p-5"><p className="text-xs font-semibold uppercase tracking-[0.14em] text-fast-primary">Rastreabilidade</p><p className="mt-3 text-sm leading-6 text-fast-text">Objetivos e atividades estão alinhados ao tema informado. As fontes e habilidades curriculares poderão ser revisadas antes da exportação.</p></div></aside>
      </div>
    </div>
  );
}

function LessonStep({ number, title, duration, text }: { number: string; title: string; duration: string; text: string }) {
  return <div className="flex gap-4"><span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-[#f6e6d9] text-xs font-bold text-fast-primary">{number}</span><div><div className="flex flex-wrap items-baseline gap-x-3 gap-y-1"><h3 className="font-semibold text-fast-heading">{title}</h3><span className="text-xs text-fast-text-muted">{duration}</span></div><p className="mt-1 text-sm leading-6 text-fast-text">{text}</p></div></div>;
}

function LessonsEmpty({ onCreate }: { onCreate: () => void }) {
  return <div className="flex min-h-[60vh] flex-col items-center justify-center text-center"><div className="flex h-14 w-14 items-center justify-center rounded-full bg-[#f6e6d9] text-fast-primary"><FolderOpen size={25} strokeWidth={1.7} aria-hidden="true" /></div><h1 className="font-display mt-6 text-4xl text-fast-heading">Suas aulas começam aqui.</h1><p className="mt-3 max-w-md text-base leading-7 text-fast-text-muted">Crie sua primeira aula e deixe o FAST CLASS ajudar a transformar a ideia em um plano aplicável.</p><button type="button" onClick={onCreate} className="focus-ring mt-7 inline-flex items-center gap-2 rounded-lg bg-fast-primary px-5 py-3 text-sm font-semibold text-white transition hover:bg-fast-primary-hover"><Plus size={18} aria-hidden="true" /> Criar primeira aula</button></div>;
}

export function FastClassWorkspace() {
  const [view, setView] = useState<View>("dashboard");
  const [mobileOpen, setMobileOpen] = useState(false);

  function navigate(nextView: View) {
    setView(nextView);
    setMobileOpen(false);
  }

  return (
    <div className="flex min-h-screen bg-fast-bg">
      <Sidebar view={view} onNavigate={navigate} />
      <div className="flex min-w-0 flex-1 flex-col">
        <MobileHeader onMenu={() => setMobileOpen((open) => !open)} />
        {mobileOpen && <div className="border-b border-fast-border-soft bg-fast-surface px-5 py-3 lg:hidden"><nav className="space-y-1" aria-label="Navegação móvel">{navItems.map(({ label, view: targetView, icon: Icon }) => <button key={targetView} type="button" onClick={() => navigate(targetView)} className="focus-ring flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left text-sm font-medium text-fast-text"><Icon size={17} aria-hidden="true" />{label}</button>)}<button type="button" onClick={() => navigate("builder")} className="focus-ring flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left text-sm font-medium text-fast-text"><Sparkles size={17} aria-hidden="true" />Criar com IA</button></nav></div>}
        <main className="mx-auto w-full max-w-[1280px] flex-1 px-5 py-8 sm:px-8 sm:py-10 lg:px-12">
          {view === "dashboard" && <Dashboard onCreate={() => navigate("builder")} />}
          {view === "builder" && <Builder onBack={() => navigate("dashboard")} onGenerated={() => navigate("result")} />}
          {view === "result" && <LessonResult onBack={() => navigate("dashboard")} />}
          {view === "lessons" && <LessonsEmpty onCreate={() => navigate("builder")} />}
        </main>
        <footer className="mx-auto w-full max-w-[1280px] px-5 pb-7 text-xs text-fast-text-muted sm:px-8 lg:px-12">FAST CLASS · seu planejamento, com mais clareza.</footer>
      </div>
    </div>
  );
}
