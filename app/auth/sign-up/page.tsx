import CreateAccountScreen from "@/components/create-account-screen";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Gerador de Aulas | Criar conta",
  description: "Crie seu workspace pessoal para planejar aulas com apoio da IA.",
};

export default function Page() {
  return <CreateAccountScreen />;
}
