"use client";

import { createClient } from "@/lib/supabase/client";
import { ArrowRight, BookOpen, Check, FileText, Layers3, ListChecks, Sparkles, Upload } from "lucide-react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useState } from "react";

import styles from "./account-type-screen.module.css";

type AccountType = "fast" | "full";

const plans: Array<{
  id: AccountType;
  name: string;
  eyebrow: string;
  description: string;
  badge: string;
  features: Array<{ label: string; icon: typeof BookOpen }>;
}> = [
  {
    id: "fast",
    name: "Fast",
    eyebrow: "Plano básico",
    description: "Para transformar uma ideia em um material utilizável rapidamente.",
    badge: "Mais rápido",
    features: [
      { label: "Texto curto, metodologia e tema", icon: BookOpen },
      { label: "Até 2 imagens e 1 arquivo de apoio", icon: Upload },
      { label: "Uma geração com pacote enxuto", icon: Sparkles },
      { label: "Rascunho editável para revisar", icon: FileText },
    ],
  },
  {
    id: "full",
    name: "Full",
    eyebrow: "Plano completo",
    description: "Para explorar uma solução pedagógica mais personalizada e completa.",
    badge: "Mais completo",
    features: [
      { label: "Contexto da turma e objetivos", icon: BookOpen },
      { label: "Vários arquivos e referências", icon: Upload },
      { label: "Planejamento, geração e revisão", icon: Layers3 },
      { label: "Alternativas, adaptações e recuperação", icon: ListChecks },
    ],
  },
];

export default function AccountTypeScreen() {
  const router = useRouter();
  const [selectedPlan, setSelectedPlan] = useState<AccountType>("fast");
  const [status, setStatus] = useState("");
  const [isSaving, setIsSaving] = useState(false);

  async function submit() {
    setIsSaving(true);
    setStatus("");

    try {
      if (process.env.NEXT_PUBLIC_SUPABASE_URL && process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY) {
        const supabase = createClient();
        const { error } = await supabase.auth.updateUser({
          data: { account_type: selectedPlan },
        });

        if (error) throw error;
      }

      window.sessionStorage.setItem("account_type", selectedPlan);
      router.push("/protected");
      router.refresh();
    } catch {
      setStatus("Não foi possível salvar sua escolha agora. Tente novamente.");
    } finally {
      setIsSaving(false);
    }
  }

  return (
    <main className={styles.screen}>
      <header className={styles.topbar}>
        <Link className={styles.lockup} href="/" aria-label="Gerador de Aulas — início">
          <span className={styles.brandMark} aria-hidden="true"><Sparkles size={17} strokeWidth={1.8} /></span>
          <span>Gerador de Aulas</span>
        </Link>
        <span className={styles.stepNote}>Etapa 2 de 2</span>
      </header>

      <section className={styles.content} aria-labelledby="account-type-title">
        <div className={styles.heading}>
          <div>
            <p className={styles.eyebrow}>CRIE SEU WORKSPACE · ÚLTIMA ETAPA</p>
            <h1 id="account-type-title">Escolha como quer começar.</h1>
            <p className={styles.intro}>
              Selecione o tipo de conta que combina com o seu momento. Você poderá mudar depois.
            </p>
          </div>
          <div className={styles.progress} aria-label="Cadastro concluído em duas etapas">
            <span className={styles.progressDone}><Check size={14} strokeWidth={2.2} /></span>
            <span className={styles.progressLine} />
            <span className={styles.progressCurrent}>2</span>
          </div>
        </div>

        <form onSubmit={(event) => { event.preventDefault(); void submit(); }}>
          <div className={styles.planGrid} role="radiogroup" aria-label="Tipo de conta">
            {plans.map((plan) => {
              const isSelected = selectedPlan === plan.id;

              return (
                <label className={`${styles.planCard} ${isSelected ? styles.selected : ""}`} key={plan.id}>
                  <input
                    type="radio"
                    name="accountType"
                    value={plan.id}
                    checked={isSelected}
                    onChange={() => setSelectedPlan(plan.id)}
                  />
                  <div className={styles.cardHeader}>
                    <span className={`${styles.planIcon} ${plan.id === "full" ? styles.fullIcon : ""}`} aria-hidden="true">
                      {plan.id === "fast" ? <Sparkles size={21} strokeWidth={1.8} /> : <Layers3 size={21} strokeWidth={1.8} />}
                    </span>
                    <span className={styles.badge}>{plan.badge}</span>
                  </div>
                  <span className={styles.planEyebrow}>{plan.eyebrow}</span>
                  <span className={styles.planName}>{plan.name}</span>
                  <span className={styles.planDescription}>{plan.description}</span>
                  <span className={styles.divider} />
                  <span className={styles.featureList}>
                    {plan.features.map(({ label, icon: FeatureIcon }) => (
                      <span className={styles.feature} key={label}>
                        <FeatureIcon size={16} strokeWidth={1.8} aria-hidden="true" />
                        <span>{label}</span>
                      </span>
                    ))}
                  </span>
                  <span className={styles.selectState} aria-hidden="true">
                    <span className={styles.radioDot}>{isSelected ? <Check size={12} strokeWidth={2.5} /> : null}</span>
                    {isSelected ? "Selecionado" : "Escolher este plano"}
                  </span>
                </label>
              );
            })}
          </div>

          <div className={styles.actions}>
            <Link className={styles.backLink} href="/auth/login">← Voltar para entrar</Link>
            <button className={styles.primary} type="submit" disabled={isSaving}>
              <span>{isSaving ? "Salvando escolha..." : `Continuar com ${selectedPlan === "fast" ? "Fast" : "Full"}`}</span>
              <ArrowRight size={18} strokeWidth={1.9} />
            </button>
          </div>
          <p className={styles.privacyNote}>
            <Check size={14} strokeWidth={2} aria-hidden="true" />
            Seu workspace começa com o plano escolhido e pode evoluir com você.
          </p>
          {status && <p className={styles.status} role="status" aria-live="polite">{status}</p>}
        </form>
      </section>
    </main>
  );
}
