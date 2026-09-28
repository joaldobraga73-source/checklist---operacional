import React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  Check,
  ChevronRight,
  CircleHelp,
  Download,
  FileDown,
  FileUp,
  Gauge,
  GripVertical,
  Info,
  Mail,
  Menu,
  Minus,
  Plus,
  Save,
  Settings2,
  Share2,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import { toast } from "sonner";
import mammoth from "mammoth/mammoth.browser";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";

type Status = "conforme" | "nao-conforme" | "nao-se-aplica" | "nao-observado" | null;
type EvidencePhoto = { src: string; comment: string; isPrimary: boolean; capturedAt: string; latitude?: number; longitude?: number };
type Question = { id: string; text: string; weight: number; status: Status; section?: string; photos?: EvidencePhoto[] };

type InspectionMeta = { responsible: string; company: string; location: string; date: string; signature: string; notes: string };
type Checklist = { title: string; code: string; questions: Question[]; meta?: InspectionMeta };
type HistoryRecord = { id: string; title: string; date: string; score: number; answered: number; total: number; company?: string; responsible?: string };

const sectionForQuestion = (text: string) => {
  const number = Number(text.match(/^(\d+)/)?.[1] || 0);
  const sections: Record<number, string> = {
    1: "1. Responsabilidade Técnica e Documentações",
    2: "2. Saúde dos Manipuladores",
    3: "3. Controle de Água da Unidade",
    4: "4. Controle de Vetores e Pragas",
    6: "6. Piso, Paredes, Teto, Portas e Janelas",
    7: "6. Piso, Paredes, Teto, Portas e Janelas",
    8: "6. Piso, Paredes, Teto, Portas e Janelas",
    9: "6. Piso, Paredes, Teto, Portas e Janelas",
    10: "6. Piso, Paredes, Teto, Portas e Janelas",
    11: "6. Piso, Paredes, Teto, Portas e Janelas",
    12: "12. Instalações Sanitárias e Vestiários",
    13: "13. Manejo de Resíduos",
    14: "14. Área de Pré-preparo e Preparo de Alimentos",
    15: "15. Higienização de Instalações, Equipamentos, Móveis e Utensílios",
    16: "16. Higiene dos Manipuladores",
    17: "17. Recebimento de Mercadoria",
    18: "18. Estoque e Armazenamento",
    19: "19. Descongelamento",
    21: "21. Preparação e Cocção",
    23: "23. Transporte de Alimentos Prontos",
    24: "24. Área de Refeitório e Alimentação de Colaboradores",
  };
  return sections[number] || "Outros itens";
};

const modelQuestions: Question[] = [
  { id: "model-1", text: "1.1 O estabelecimento possui responsável técnico (gerente/ líder) capacitado c/ certificado de curso que aborde tópicos essenciais de segurança alimentar, Contaminantes alimentares, Doenças transmitidas por alimentos (DTAs) e Boas Práticas de Manipulação de Alimentos?", weight: 1.16, status: null },
  { id: "model-2", text: "1.2 Possui Manual de Boas Práticas e POP atualizados?", weight: 1.16, status: null },
  { id: "model-3", text: "1.3 Possui Alvará de Funcionamento e Sanitário atualizado?", weight: 1.16, status: null },
  { id: "model-4", text: "1.4 Possui PGR (Programa Gerenciamento de Riscos) atualizado?", weight: 1.16, status: null },
  { id: "model-5", text: "1.5 A Unidade possui PCMSO (Programa de Controle Médico de Saúde Ocupacional) ou ASO anual, exames, atualizados?", weight: 1.16, status: null },
  { id: "model-6", text: "2.1 Os colaboradores que apresentam feridas, lesões, gripe ou outra situação de risco, são afastados da manipulação de alimentos e recebem orientações de como proceder?", weight: 1.16, status: null },
  { id: "model-7", text: "3.1 A água usada no preparo dos alimentos é tratada e de sistema público?", weight: 1.16, status: null },
  { id: "model-8", text: "3.5 Possui reservatório tampado e em bom estado de conservação?", weight: 1.16, status: null },
  { id: "model-9", text: "3.7 A higienização é feita a cada seis meses por empresa especializada e com registro?", weight: 1.16, status: null },
  { id: "model-10", text: "3.8 Utiliza-se gelo na Unidade, é fabricado com água filtrada?", weight: 1.16, status: null },
  { id: "model-11", text: "4.1 O controle de pragas é realizado controle químico por dedetizadora terceirizada?", weight: 1.16, status: null },
  { id: "model-12", text: "4.2 Existe registro, certificado, OS da última visita técnica de dedetização?", weight: 1.16, status: null },
  { id: "model-13", text: "4.3 Os arredores da Unidade apresentam ausência de animais domésticos?", weight: 1.16, status: null },
  { id: "model-14", text: "4.5 O lixo é devidamente tampado e apresenta ausência de lixo nos arredores?", weight: 1.16, status: null },
  { id: "model-15", text: "4.7 Os dispositivos, caixas de iscas estão devidamente posicionados e em bom estado de conservação?", weight: 1.16, status: null },
  { id: "model-16", text: "6.1 O piso da Unidade está limpo?", weight: 1.16, status: null },
  { id: "model-17", text: "6.2 Os ralos da cozinha estão limpos, desobstruídos e funcionando?", weight: 1.16, status: null },
  { id: "model-18", text: "7.1 As paredes da Unidade são de material liso, empermeável, lavável, em bom estado de conservação e limpeza?", weight: 1.16, status: null },
  { id: "model-19", text: "8.1 O teto e forro da Unidade são mantidos íntegros, conservados e limpos, livre de bolores e teias de aranha?", weight: 1.16, status: null },
  { id: "model-20", text: "9.1 As portas, janelas e telas da Unidade estão limpas?", weight: 1.16, status: null },
  { id: "model-21", text: "10.3 Dentro da área de produção as lâmpadas e protetores estão limpos?", weight: 1.16, status: null },
  { id: "model-22", text: "11.5 Possui ventiladores ligados na cozinha, com corrente de ar de forma a não incidir sobre os alimentos nas áreas de preparo?", weight: 1.16, status: null },
  { id: "model-23", text: "12.4 Os vasos sanitários de colaboradores estão tampados e limpos?", weight: 1.16, status: null },
  { id: "model-24", text: "12.5 Nos sanitários: os facilitadores (toalha de papel descartável e não reciclável, sabonete líquido inodoro, antisséptico) estão abastecidos?", weight: 1.16, status: null },
  { id: "model-25", text: "12.7 As lixeiras dos sanitários estão limpas e com capacidades adequadas, sem excesso de lixo?", weight: 1.16, status: null },
  { id: "model-26", text: "12.8 Nos sanitários há presença de avisos para higiene das mãos?", weight: 1.16, status: null },
  { id: "model-27", text: "12.9 Os armários de colaboradores estão fechados e os pertences pessoais guardados?", weight: 1.16, status: null },
  { id: "model-28", text: "13.2 O depósito de lixo externo é isolado, tampado ou tratado de forma a evitar contaminação?", weight: 1.16, status: null },
  { id: "model-29", text: "13.3 Os cestos de lixo estão limpos, tampados em número compatível com a produção, sem excessos?", weight: 1.16, status: null },
  { id: "model-30", text: "13.7 Há horários programados para retirada do lixo ou ao final do expediente?", weight: 1.16, status: null },
  { id: "model-31", text: "13.8 Os colaboradores usam luvas de limpeza para retirar o lixo e higienizam a mão ao retornar para cozinha?", weight: 1.16, status: null },
  { id: "model-32", text: "13.11 As caixas de gordura são limpas a cada 3 meses e possuem registro?", weight: 1.16, status: null },
  { id: "model-33", text: "13.12 Dentro da UAN os resíduos orgânicos são separados dos secos (coleta seletiva)?", weight: 1.16, status: null },
  { id: "model-34", text: "14.1 Possui pia exclusiva para pré-preparo de hortifrutis? É realizada a higienização completa dos hortifrútis, lavagem, desinfecção e enxague (com sanitizante, cloro)?", weight: 1.16, status: null },
  { id: "model-35", text: "14.3 Na área de produção as bancadas e pias são lisas, de material impermeável, lavável, estão limpas organizadas e conservadas?", weight: 1.16, status: null },
  { id: "model-36", text: "14.4 Na área de produção os equipamentos estão limpos organizados e conservados?", weight: 1.16, status: null },
  { id: "model-37", text: "14.2 É realizada higienização das embalagens antes do preparo?", weight: 1.16, status: null },
  { id: "model-38", text: "14.5 A área produção possui utensílios suficientes e de acordo com as preparações, em bom estado de conservação e limpeza?", weight: 1.16, status: null },
  { id: "model-39", text: "14.7 A área de produção é livre de materiais em desuso e madeira?", weight: 1.16, status: null },
  { id: "model-40", text: "14.9 Existe termômetro com calibração vigente para monitorar a temperatura do alimento durante a preparação?", weight: 1.16, status: null },
  { id: "model-41", text: "14.11 Existe planilha diária preenchida de controle de temperatura?", weight: 1.16, status: null },
  { id: "model-42", text: "14.13 Existe pia exclusiva e desobstruída, dotada de papel, sabonete líquido, álcool 70% e lixeira com tampa e acionamento por pedal para higienização das mãos dentro da produção?", weight: 1.16, status: null },
  { id: "model-43", text: "14 Os manipuladores de alimentos higienizam as mãos frequentemente e troca de luvas sempre que necessário?", weight: 1.16, status: null },
  { id: "model-44", text: "14.18 Na cozinha a disposição de armazenamento segue PEPS (primeiro que entra, primeiro que sai)?", weight: 1.16, status: null },
  { id: "model-45", text: "14.19 Os produtos em embalagens danificadas, estragados ou vencidos são armazenados em local separado e identificado para descarte?", weight: 1.16, status: null },
  { id: "model-46", text: "14.20 Os alimentos armazenados transferidos de embalagem, abertos e fracionados, estão tampados, protegidos e identificados com prazo de validade?", weight: 1.16, status: null },
  { id: "model-47", text: "14.23 Todo produto armazenado está dentro do prazo de validade?", weight: 1.16, status: null },
  { id: "model-48", text: "14.24 A área de produção está livre de vestígios de moscas, insetos, pragas e roedores?", weight: 1.16, status: null },
  { id: "model-49", text: "14.26 Na área de produção não há degustação de alimentos?", weight: 1.16, status: null },
  { id: "model-50", text: "14.27 A área de manipulação está livre de utilização de celular e carregadores de celular?", weight: 1.16, status: null },
  { id: "model-51", text: "14.28 Os produtos estocados na cozinha ou em manipulação estão sob paletes (15cm), sem contato direto com o chão?", weight: 1.16, status: null },
  { id: "model-52", text: "14 As geladeiras e freezes da cozinha estão organizados, limpos, sem excesso de gelo, sem presença de materiais pessoais e os alimentos separados por categoria?", weight: 1.16, status: null },
  { id: "model-53", text: "15.1 Na área de produção possui borrifadores de solução clorada, água sanitária para desinfecção de ambiente, alimentos e equipamentos?", weight: 1.16, status: null },
  { id: "model-54", text: "15.2 Existe registro (POP) de limpeza preenchido corretamente?", weight: 1.16, status: null },
  { id: "model-55", text: "15.5 A Unidade está livre de produtos de limpeza, spray, aerossóis, odorizantes ou desodorantes?", weight: 1.16, status: null },
  { id: "model-56", text: "15.7 A área de produção está livre de materiais de limpeza, vassoura, rodos, panos e similares?", weight: 1.16, status: null },
  { id: "model-57", text: "15.8 Os utensílios utilizados estão limpos e são devidamente higienizados, sem etiquetas antigas?", weight: 1.16, status: null },
  { id: "model-58", text: "15.11 Os armários de massas são varridos diariamente ao final do expediente e retirados o excesso de resíduos?", weight: 1.16, status: null },
  { id: "model-59", text: "15.12 As pás de madeira são raspadas diariamente, estão limpas sem presença de resíduos e em bom estado de conservação?", weight: 1.16, status: null },
  { id: "model-60", text: "15.13 A coifa está limpa, sem gotejamento e sujeiras aparentes?", weight: 1.16, status: null },
  { id: "model-61", text: "15.14 A pia de higienização de utensílios está sem acúmulo excessivo, o volume de louça suja está compatível com o espaço?", weight: 1.16, status: null },
  { id: "model-62", text: "16.3 Os colaboradores estão uniformizados, com uniformes limpos e em bom estado de conservação?", weight: 1.16, status: null },
  { id: "model-63", text: "16.4 Todos que acessam a área de produção, manipuladores e visitantes usam toucas ou boné para proteção dos cabelos?", weight: 1.16, status: null },
  { id: "model-64", text: "16.5 Os manipuladores estão sem adornos (anéis, brincos, relógios, pulseiras, correntes ou similares)?", weight: 1.16, status: null },
  { id: "model-65", text: "16.7 Todos os colaboradores usam EPIs, estão de sapatos de segurança, fechados e antiderrapantes, avental, luvas anticorte?", weight: 1.16, status: null },
  { id: "model-66", text: "16.9 Manipuladores mantem boa higiene pessoal, cabelo e barba aparadas, unhas curtas e sem esmalte?", weight: 1.16, status: null },
  { id: "model-67", text: "17.2 No recebimento é conferido temperatura, quantidade, qualidade, datas de validade, condições das embalagens, avaliação sensorial (cor, cheiro e aparência) dos produtos?", weight: 1.16, status: null },
  { id: "model-68", text: "18.1 No estoque os alimentos são armazenados de acordo com suas especificações, separados por tipo perecível e não perecível como congelado, refrigerado e seco?", weight: 1.16, status: null },
  { id: "model-69", text: "18.2 No estoque os produtos estão dispostos como PEPS (primeiro que entra, primeiro que sai)?", weight: 1.16, status: null },
  { id: "model-70", text: "18.3 As áreas de armazenamentos e estoque estão em bom estado de conservação, prateleiras limpas, livre de materiais em desuso e excesso de caixas de papelão?", weight: 1.16, status: null },
  { id: "model-71", text: "18.4 A câmara fria ou refrigerador está organizado (alimentos prontos pra consumo, cozidos separados dos crus, carnes separadas de hortifrutis?", weight: 1.16, status: null },
  { id: "model-72", text: "18.6 No estoque os produtos de limpeza e descartáveis estão separados dos alimentos?", weight: 1.16, status: null },
  { id: "model-73", text: "18.9 No estoque e câmaras os produtos em embalagens danificadas, estragados ou vencidos são armazenados em local separado e identificado para descarte?", weight: 1.16, status: null },
  { id: "model-74", text: "18.10 No estoque e câmaras os alimentos armazenados transferidos de embalagem, abertos e fracionados estão rotulados, identificados, datados com prazo de validade vigente e protegidos, sem alimentos vencidos?", weight: 1.16, status: null },
  { id: "model-75", text: "18.13 Existe planilha preenchida de monitoramento de temperatura das câmaras e refrigeradores da área de estoque e armazenamento?", weight: 1.16, status: null },
  { id: "model-76", text: "18.17 A área de estoque e câmaras estão livres de vestigios de pragas, moscas, insetos e roedores?", weight: 1.16, status: null },
  { id: "model-77", text: "18.20 A área de armazenamento e câmaras está livre de degustação de alimentos?", weight: 1.16, status: null },
  { id: "model-78", text: "18.21 No estoque e câmaras, todos os produtos estão sob paletes (15cm do piso), sem contato direto com o chão, a 5cm da parede e 45cm do teto?", weight: 1.16, status: null },
  { id: "model-79", text: "19.1 O descongelamento é feito sob refrigeração, quando possível, protegido, identificados com etiquetas de descongelamento e prazo de até 3 dias para descongelamento?", weight: 1.16, status: null },
  { id: "model-80", text: "21.2 Os óleos e gorduras das frituras são aquecidos no máximo até 180ºC, sendo observado o ponto de saturação, mudança de cor, presença de espuma e fumaça são descartados?", weight: 1.16, status: null },
  { id: "model-81", text: "21.4 O óleo das frituras quando desprezado obedece a regras de sustentabilidade, não sendo descartados no meio ambiente, são armazenados em recipientes com tampas, em local apropriado enquanto aguardam o recolhimento?", weight: 1.16, status: null },
  { id: "model-82", text: "23.1 No transporte de mercadorias e alimentos prontos o veículo apresenta proteção de carga?", weight: 1.16, status: null },
  { id: "model-83", text: "23.2 O veículo de transporte é revestido de material lavável, refrigerado, em bom estado de conservação e limpeza?", weight: 1.16, status: null },
  { id: "model-84", text: "24.1 No refeitório de colaboradores a área está limpa e organizada, não havendo desperdício de alimentos?", weight: 1.16, status: null },
  { id: "model-85", text: "24.6 O bebedouro de colaboradores está em bom estado de conservação e limpeza, os filtros são trocados há cada seis meses com registro de troca?", weight: 1.16, status: null },
  { id: "model-86", text: "24.8 As lixeiras do bebedouro e área do refeitório de colaboradores estão tampadas com capacidade adequada, sem excesso de lixo?", weight: 1.4, status: null }
].map((question) => ({ ...question, section: sectionForQuestion(question.text) }));

const starter: Checklist = {
  title: "Checklist de Boas Práticas",
  code: "CK-BOAS-PRATICAS",
  questions: modelQuestions,
  meta: { responsible: "", company: "", location: "", date: new Date().toISOString().slice(0, 10), signature: "", notes: "" },
};

const uid = () => `q-${Date.now()}-${Math.random().toString(16).slice(2)}`;
const normalize = (value: string) => Number(value.replace(",", ".")) || 0;
const normalizeChecklist = (value: any): Checklist => ({ ...value, questions: (value.questions || []).map((question: any) => { const text = String(question.text || "").replace(/^1\.5 É realizada higienização das embalagens/, "14.2 É realizada higienização das embalagens").replace(/^14\.29 Na área de produção possui borrifadores/, "15.1 Na área de produção possui borrifadores"); return { ...question, text, section: sectionForQuestion(text), photos: (question.photos || []).map((photo: any) => typeof photo === "string" ? { src: photo, comment: "", isPrimary: false, capturedAt: new Date().toISOString() } : photo) }; }) });
const hashPassword = async (value: string) => { const data = new TextEncoder().encode(value); const digest = await crypto.subtle.digest("SHA-256", data); return Array.from(new Uint8Array(digest)).map((item) => item.toString(16).padStart(2, "0")).join(""); };

function SignaturePad({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const drawing = useRef(false);
  const point = (event: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current; if (!canvas) return null;
    const rect = canvas.getBoundingClientRect();
    return { x: (event.clientX - rect.left) * (canvas.width / rect.width), y: (event.clientY - rect.top) * (canvas.height / rect.height) };
  };
  const start = (event: React.PointerEvent<HTMLCanvasElement>) => { const ctx = canvasRef.current?.getContext("2d"); const p = point(event); if (!ctx || !p) return; drawing.current = true; ctx.beginPath(); ctx.moveTo(p.x, p.y); };
  const move = (event: React.PointerEvent<HTMLCanvasElement>) => { if (!drawing.current) return; const ctx = canvasRef.current?.getContext("2d"); const p = point(event); if (!ctx || !p) return; ctx.lineWidth = 2.2; ctx.lineCap = "round"; ctx.strokeStyle = "#17252b"; ctx.lineTo(p.x, p.y); ctx.stroke(); };
  const finish = () => { if (!drawing.current) return; drawing.current = false; const canvas = canvasRef.current; if (canvas) onChange(canvas.toDataURL("image/png")); };
  const clear = () => { const canvas = canvasRef.current; const ctx = canvas?.getContext("2d"); if (canvas && ctx) { ctx.clearRect(0, 0, canvas.width, canvas.height); onChange(""); } };
  return <div className="signature-wrap"><div className="signature-toolbar"><span>Assine com o dedo</span><button type="button" onClick={clear}>Limpar</button></div><canvas ref={canvasRef} width={700} height={140} className="signature-canvas" onPointerDown={start} onPointerMove={move} onPointerUp={finish} onPointerLeave={finish} />{value && <span className="signature-saved">Assinatura registrada nesta inspeção</span>}</div>;
}

export default function Home() {
  const [checklist, setChecklist] = useState<Checklist>(() => {
    try {
      return normalizeChecklist(JSON.parse(localStorage.getItem("cherry-checklist") || "null") || starter);
    } catch {
      return starter;
    }
  });
  const [history, setHistory] = useState<HistoryRecord[]>(() => {
    try { return JSON.parse(localStorage.getItem("cherry-history") || "[]"); } catch { return []; }
  });
  const [profile, setProfile] = useState(() => {
    try { return JSON.parse(localStorage.getItem("cherry-profile") || '{"name":"","email":""}'); } catch { return { name: "", email: "" }; }
  });
  const [activeTab, setActiveTab] = useState<"checklist" | "estrutura">("checklist");
  const [showMenu, setShowMenu] = useState(false);
  const [showHelp, setShowHelp] = useState(false);
  const [installEvent, setInstallEvent] = useState<any>(null);
  const [filterCompany, setFilterCompany] = useState("");
  const [filterResponsible, setFilterResponsible] = useState("");
  const [filterFrom, setFilterFrom] = useState("");
  const [filterTo, setFilterTo] = useState("");
  const [authReady, setAuthReady] = useState(false);
  const [localAccount, setLocalAccount] = useState<{ phone: string; passwordHash: string } | null>(null);
  const [authMode, setAuthMode] = useState<"create" | "unlock">("unlock");
  const [authPhone, setAuthPhone] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authPasswordConfirm, setAuthPasswordConfirm] = useState("");
  const [authError, setAuthError] = useState("");
  const [unlocked, setUnlocked] = useState(false);
  const [showPasswordChange, setShowPasswordChange] = useState(false);
  const [newPassword, setNewPassword] = useState("");
  const [newPasswordConfirm, setNewPasswordConfirm] = useState("");

  useEffect(() => {
    try { const saved = JSON.parse(localStorage.getItem("cherry-local-account") || "null"); setLocalAccount(saved); setAuthMode(saved ? "unlock" : "create"); setAuthPhone(saved?.phone || ""); } catch { setAuthMode("create"); } finally { setAuthReady(true); }
  }, []);

  const submitAuth = async (event: React.FormEvent) => {
    event.preventDefault(); setAuthError("");
    if (authPassword.length < 6) { setAuthError("A senha precisa ter pelo menos 6 caracteres."); return; }
    if (authMode === "create") {
      if (!authPhone.trim()) { setAuthError("Informe seu telefone."); return; }
      if (authPassword !== authPasswordConfirm) { setAuthError("As senhas não coincidem."); return; }
      const account = { phone: authPhone.trim(), passwordHash: await hashPassword(authPassword) }; localStorage.setItem("cherry-local-account", JSON.stringify(account)); setLocalAccount(account); setUnlocked(true); setAuthPassword(""); setAuthPasswordConfirm(""); toast.success("Acesso local criado neste celular.");
    } else {
      if (!localAccount || await hashPassword(authPassword) !== localAccount.passwordHash) { setAuthError("Telefone ou senha incorretos."); return; }
      setUnlocked(true); setAuthPassword("");
    }
  };

  const changePassword = async () => {
    if (newPassword.length < 6) { toast.error("A nova senha precisa ter pelo menos 6 caracteres."); return; }
    if (newPassword !== newPasswordConfirm) { toast.error("As novas senhas não coincidem."); return; }
    const account = { ...(localAccount || { phone: authPhone }), passwordHash: await hashPassword(newPassword) };
    localStorage.setItem("cherry-local-account", JSON.stringify(account)); setLocalAccount(account); setNewPassword(""); setNewPasswordConfirm(""); setShowPasswordChange(false); toast.success("Senha alterada neste dispositivo.");
  };

  const exportHistoryCsv = () => {
    const rows = [["Data", "Checklist", "Empresa", "Responsável", "Resultado (%)", "Respondidas", "Total"], ...history.map((item) => [new Date(item.date).toLocaleString("pt-BR"), item.title, item.company || "", item.responsible || "", String(item.score), String(item.answered), String(item.total)])];
    const csv = "\uFEFF" + rows.map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(";")).join("\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" })); const a = document.createElement("a"); a.href = url; a.download = "cherry-historico-inspecoes.csv"; a.click(); URL.revokeObjectURL(url); toast.success("Histórico exportado para Excel/CSV.");
  };

  const exportHistoryPdf = () => {
    const popup = window.open("", "_blank", "width=900,height=700"); if (!popup) { toast.error("Permita pop-ups para gerar o PDF do histórico."); return; }
    const rows = history.map((item) => `<tr><td>${new Date(item.date).toLocaleString("pt-BR")}</td><td>${item.title}</td><td>${item.company || "—"}</td><td>${item.responsible || "—"}</td><td><strong>${item.score}%</strong></td><td>${item.answered}/${item.total}</td></tr>`).join("");
    popup.document.write(`<html><head><title>Cherry Checklist — Histórico</title><style>body{font-family:Arial,sans-serif;color:#17252b;padding:30px}h1{color:#b51f42}p{color:#728087}table{width:100%;border-collapse:collapse;margin-top:22px}th,td{text-align:left;padding:10px;border-bottom:1px solid #dfe9e5;font-size:12px}th{background:#f2f7f5;color:#17775c}strong{color:#b51f42}</style></head><body><h1>🍒 Cherry Checklist</h1><h2>Histórico de inspeções</h2><p>Relatório gerado em ${new Date().toLocaleString("pt-BR")}</p><table><thead><tr><th>Data</th><th>Checklist</th><th>Empresa</th><th>Responsável</th><th>Resultado</th><th>Itens</th></tr></thead><tbody>${rows || "<tr><td colspan=6>Nenhuma inspeção salva.</td></tr>"}</tbody></table><script>window.onload=()=>window.print()<\/script></body></html>`); popup.document.close();
  };

  const totalWeight = useMemo(() => checklist.questions.reduce((sum, q) => sum + q.weight, 0), [checklist.questions]);
  const formatPercent = (value: number) => Math.abs(value - 100) < 0.01 ? "100" : Number(value.toFixed(2)).toString();
  const answered = checklist.questions.filter((q) => q.status !== null).length;
  const applicableWeight = useMemo(() => checklist.questions.reduce((sum, q) => sum + (q.status !== "nao-se-aplica" ? q.weight : 0), 0), [checklist.questions]);
  const score = useMemo(() => { const points = checklist.questions.reduce((sum, q) => sum + (q.status === "conforme" ? q.weight : 0), 0); return applicableWeight ? Math.round((points / applicableWeight) * 100) : 0; }, [checklist.questions, applicableWeight]);
  const progress = checklist.questions.length ? Math.round((answered / checklist.questions.length) * 100) : 0;
  const isValid = Math.abs(totalWeight - 100) < 0.01;

  useEffect(() => {
    localStorage.setItem("cherry-checklist", JSON.stringify(checklist));
  }, [checklist]);

  useEffect(() => { localStorage.setItem("cherry-history", JSON.stringify(history)); }, [history]);
  useEffect(() => { localStorage.setItem("cherry-profile", JSON.stringify(profile)); }, [profile]);
  useEffect(() => { const handler = (event: Event) => { event.preventDefault(); setInstallEvent(event); }; window.addEventListener("beforeinstallprompt", handler); return () => window.removeEventListener("beforeinstallprompt", handler); }, []);
  const installApp = async () => { if (installEvent) { installEvent.prompt(); await installEvent.userChoice; setInstallEvent(null); } else { toast.info("No iPhone: toque em Compartilhar e depois em Adicionar à Tela de Início. No Android, use o menu do navegador e escolha Instalar app."); } };

  const updateQuestion = (id: string, update: Partial<Question>) =>
    setChecklist((current) => ({ ...current, questions: current.questions.map((q) => (q.id === id ? { ...q, ...update } : q)) }));

  const addQuestion = () => {
    setChecklist((current) => ({
      ...current,
      questions: [...current.questions, { id: uid(), text: "Nova pergunta", weight: 0, status: null }],
    }));
    setActiveTab("estrutura");
  };

  const loadModelChecklist = () => {
    setChecklist({ ...starter, questions: modelQuestions.map((question) => ({ ...question, id: uid() })), meta: { ...starter.meta!, date: new Date().toISOString().slice(0, 10) } });
    setActiveTab("estrutura");
    toast.success("Modelo de Boas Práticas carregado com 86 perguntas.");
  };

  const removeQuestion = (id: string) => setChecklist((current) => ({ ...current, questions: current.questions.filter((q) => q.id !== id) }));

  const compressPhoto = (file: File) => new Promise<string>((resolve, reject) => { const image = new Image(); const reader = new FileReader(); reader.onload = () => { image.onload = () => { const scale = Math.min(1, 1200 / image.width); const canvas = document.createElement("canvas"); canvas.width = Math.round(image.width * scale); canvas.height = Math.round(image.height * scale); canvas.getContext("2d")?.drawImage(image, 0, 0, canvas.width, canvas.height); resolve(canvas.toDataURL("image/jpeg", 0.72)); }; image.onerror = reject; image.src = String(reader.result); }; reader.onerror = reject; reader.readAsDataURL(file); });
  const addPhotos = async (id: string, files: FileList | null) => { if (!files?.length) return; try { const capturedAt = new Date().toISOString(); const sources = await Promise.all(Array.from(files).slice(0, 4).map(compressPhoto)); const photos: EvidencePhoto[] = sources.map((src) => ({ src, comment: "", isPrimary: false, capturedAt })); setChecklist((current) => ({ ...current, questions: current.questions.map((question) => question.id === id ? { ...question, photos: [...(question.photos || []), ...photos].slice(0, 4) } : question) })); toast.success("Foto anexada à avaliação."); } catch { toast.error("Não foi possível adicionar esta foto."); } };
  const removePhoto = (id: string, index: number) => setChecklist((current) => ({ ...current, questions: current.questions.map((question) => question.id === id ? { ...question, photos: (question.photos || []).filter((_, photoIndex) => photoIndex !== index) } : question) }));
  const updatePhoto = (id: string, index: number, update: Partial<EvidencePhoto>) => setChecklist((current) => ({ ...current, questions: current.questions.map((question) => question.id === id ? { ...question, photos: (question.photos || []).map((photo, photoIndex) => photoIndex === index ? { ...photo, ...update } : update.isPrimary ? { ...photo, isPrimary: false } : photo) } : question) }));

  const save = () => {
    if (!isValid) {
      toast.error(`Os pesos precisam somar 100%. Atualmente somam ${formatPercent(totalWeight)}%.`);
      return;
    }
    const record: HistoryRecord = { id: uid(), title: checklist.title, date: new Date().toISOString(), score, answered, total: checklist.questions.length, company: checklist.meta?.company, responsible: checklist.meta?.responsible };
    setHistory((current) => [record, ...current.filter((item) => item.title !== checklist.title)].slice(0, 8));
    toast.success("Checklist salvo no histórico deste dispositivo.");
  };

  const filteredHistory = history.filter((item) => {
    const day = item.date.slice(0, 10);
    return (!filterCompany || (item.company || "").toLowerCase().includes(filterCompany.toLowerCase())) && (!filterResponsible || (item.responsible || "").toLowerCase().includes(filterResponsible.toLowerCase())) && (!filterFrom || day >= filterFrom) && (!filterTo || day <= filterTo);
  });

  const exportJson = () => {
    const blob = new Blob([JSON.stringify(checklist, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${checklist.title.toLowerCase().replace(/\s+/g, "-")}.json`;
    a.click();
    URL.revokeObjectURL(url);
    toast.success("Estrutura exportada.");
  };

  const exportQuestionsCsv = () => {
    const rows = [["Pergunta", "Peso (%)"], ...checklist.questions.map((question) => [question.text, String(question.weight)])];
    const csv = "\uFEFF" + rows.map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(";")).join("\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" })); const a = document.createElement("a"); a.href = url; a.download = "cherry-perguntas.csv"; a.click(); URL.revokeObjectURL(url); toast.success("Perguntas exportadas em CSV.");
  };

  const exportQuestionsWord = () => {
    const rows = checklist.questions.map((question, index) => `<p><strong>${index + 1}. ${escapeHtml(question.text)}</strong><br>Peso: ${question.weight}%</p>`).join("");
    const content = `<!doctype html><html><head><meta charset="utf-8"><title>Perguntas - ${escapeHtml(checklist.title)}</title></head><body><h1>🍒 ${escapeHtml(checklist.title)}</h1><p>Lista de perguntas para o checklist</p>${rows}</body></html>`;
    const url = URL.createObjectURL(new Blob([content], { type: "application/msword" })); const a = document.createElement("a"); a.href = url; a.download = "cherry-perguntas.doc"; a.click(); URL.revokeObjectURL(url); toast.success("Perguntas exportadas para abrir no Word.");
  };

  const downloadWordTemplate = () => { const a = document.createElement("a"); a.href = "/manus-storage/Modelo_Checklist_Boas_Praticas_cc3537f7.docx"; a.download = "CheckList_VISITA_TECNICA_Boas_Praticas.docx"; a.target = "_blank"; a.click(); toast.success("Modelo Word aberto para download."); };

  const downloadQuestionTemplate = () => {
    const csv = "\uFEFFPergunta;Peso (%)\nOs equipamentos estão em boas condições?;25\nA área está limpa e organizada?;25\nOs documentos estão atualizados?;25\nOs procedimentos foram seguidos?;25\n";
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" })); const a = document.createElement("a"); a.href = url; a.download = "modelo-perguntas-cherry.csv"; a.click(); URL.revokeObjectURL(url); toast.success("Modelo baixado. Preencha e importe no app.");
  };

  const importQuestionsFile = async (file?: File) => {
    if (!file) return;
    if (/\.pdf$/i.test(file.name)) {
      try {
        const pdf = await (pdfjsLib.getDocument as any)({ data: await file.arrayBuffer(), disableWorker: true }).promise;
        const pages: string[] = [];
        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) { const page = await pdf.getPage(pageNumber); const content = await page.getTextContent(); pages.push(content.items.map((item: any) => "str" in item ? item.str : "").join(" ")); }
        const rows = pages.join("\n").split(/\r?\n/).flatMap((line) => line.split(/(?<=\?)\s+(?=[A-ZÁÉÍÓÚÃÕ])/)).map((line) => line.replace(/^\s*(?:[•*-]|\d+[.)])\s*/, "").trim()).filter(Boolean);
        if (!rows.length) throw new Error();
        setChecklist((current) => ({ ...current, questions: rows.map((text) => ({ id: uid(), text, weight: 0, status: null })) })); setActiveTab("estrutura"); toast.success(`${rows.length} perguntas importadas do PDF. Confira os pesos antes de salvar.`);
      } catch { toast.error("Não foi possível ler este PDF. Use um PDF com texto selecionável, não apenas imagem."); }
      return;
    }
    if (/\.(xlsx|xls)$/i.test(file.name)) {
      try {
        const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const sheetRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
        const first = Array.isArray(sheetRows[0]) ? sheetRows[0] : [];
        const hasHeader = /pergunta|question|item/i.test(String(first[0] || ""));
        const rows = sheetRows.slice(hasHeader ? 1 : 0).map((row) => { const values = Array.isArray(row) ? row : []; return { text: String(values[0] || "").trim(), weight: normalize(String(values[1] || "0")) }; }).filter((row) => row.text && !/^pergunta$/i.test(row.text));
        if (!rows.length) throw new Error();
        setChecklist((current) => ({ ...current, questions: rows.map((row) => ({ id: uid(), text: row.text, weight: row.weight, status: null })) }));
        setActiveTab("estrutura");
        toast.success(`${rows.length} perguntas importadas do Excel. Confira os pesos antes de salvar.`);
      } catch { toast.error("Não foi possível ler este Excel. Use .xlsx ou .xls com a pergunta na primeira coluna e o peso na segunda."); }
      return;
    }
    if (file.name.toLowerCase().endsWith(".doc")) { toast.error("O Word antigo .doc não é lido diretamente. Salve o arquivo como .docx no Word e importe novamente."); return; }
    if (file.name.toLowerCase().endsWith(".docx")) {
      try {
        const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
        const rows = result.value.split(/\r?\n/).map((line) => line.replace(/^\s*(?:[•*-]|\d+(?:\.\d+)?[.)]?)\s*/, "").trim()).filter((line) => line.includes("?") && line.length > 8);
        if (!rows.length) throw new Error();
        setChecklist((current) => ({ ...current, questions: rows.map((text) => ({ id: uid(), text, weight: 0, status: null })) }));
        setActiveTab("estrutura");
        toast.success(`${rows.length} perguntas importadas do Word. Confira os pesos antes de salvar.`);
      } catch { toast.error("Não foi possível ler o Word. Use um arquivo .docx com uma pergunta por linha."); }
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const raw = String(reader.result || "").replace(/^\uFEFF/, "");
        if (file.name.toLowerCase().endsWith(".json")) {
          const imported = JSON.parse(raw) as Checklist;
          if (!imported.questions || !Array.isArray(imported.questions)) throw new Error();
          setChecklist(normalizeChecklist(imported)); toast.success("Checklist JSON importado."); return;
        }
        const lines = raw.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
        if (lines.length < 2) throw new Error();
        const delimiter = lines[0].includes(";") ? ";" : ",";
        const parseLine = (line: string) => { const values: string[] = []; let value = ""; let quoted = false; for (let i = 0; i < line.length; i += 1) { const char = line[i]; if (char === '"' && line[i + 1] === '"') { value += '"'; i += 1; } else if (char === '"') quoted = !quoted; else if (char === delimiter && !quoted) { values.push(value.trim()); value = ""; } else value += char; } values.push(value.trim()); return values; };
        const rows = lines.slice(1).map(parseLine).map((row) => ({ text: row[0], weight: normalize(row[1] || "0") })).filter((row) => row.text);
        if (!rows.length) throw new Error();
        setChecklist((current) => ({ ...current, questions: rows.map((row) => ({ id: uid(), text: row.text, weight: row.weight, status: null })) })); setActiveTab("estrutura"); toast.success(`${rows.length} perguntas importadas. Confira os pesos antes de salvar.`);
      } catch { toast.error("Arquivo inválido. Use o modelo CSV do Cherry ou um JSON exportado pelo app."); }
    };
    reader.readAsText(file);
  };

  const reportFileName = () => `${checklist.title.toLowerCase().replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "cherry-checklist"}-relatorio.html`;
  const escapeHtml = (value: string) => value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[char] || char));
  const reportHtml = () => {
    const meta = checklist.meta || starter.meta!;
    const rows = checklist.questions.map((question, index) => `${(index === 0 || question.section !== checklist.questions[index - 1]?.section) ? `<tr class="section-row"><td colspan="5">${escapeHtml(question.section || "Outros itens")}</td></tr>` : ""}<tr><td>${index + 1}</td><td>${escapeHtml(question.text)}</td><td>${question.weight}%</td><td class="${question.status === "conforme" ? "ok" : question.status === "nao-conforme" ? "bad" : "pending"}">${question.status === "conforme" ? "CONFORME" : question.status === "nao-conforme" ? "NÃO CONFORME" : question.status === "nao-se-aplica" ? "NÃO SE APLICA" : question.status === "nao-observado" ? "NÃO OBSERVADO" : "NÃO RESPONDIDO"}</td><td>${question.status === "conforme" ? question.weight : 0}%</td></tr>${(question.photos || []).map((photo) => `<tr><td></td><td colspan="4"><img src="${photo.src}" style="max-width:220px;max-height:140px;border-radius:6px;margin:4px" /><br><small>${photo.isPrimary ? "FOTO PRINCIPAL · " : ""}${escapeHtml(photo.comment || "Sem comentário")} · ${new Date(photo.capturedAt).toLocaleString("pt-BR")}${photo.latitude ? ` · GPS: ${photo.latitude.toFixed(5)}, ${photo.longitude?.toFixed(5)}` : ""}</small></td></tr>`).join("")}`).join("");
    return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>${escapeHtml(checklist.title)} — Relatório</title><style>body{font-family:Arial,sans-serif;color:#17252b;max-width:1000px;margin:0 auto;padding:32px}h1{color:#b51f42;margin:0 0 5px}h2{margin-top:30px;color:#17775c}p{color:#697a78}table{width:100%;border-collapse:collapse;margin-top:15px}th,td{text-align:left;padding:10px 8px;border-bottom:1px solid #dce7e3;font-size:12px}th{background:#f0f7f4;color:#17775c}.section-row td{background:#fff0f3;color:#b51f42;font-weight:700;font-size:13px;padding-top:16px}.ok{color:#17775c;font-weight:700}.bad{color:#b51f42;font-weight:700}.pending{color:#8c9895}.summary{display:flex;gap:30px;padding:18px;background:#f4f8f6;border-radius:10px}.summary b{font-size:24px;color:#b51f42}.meta{display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px}.notes{white-space:pre-wrap;border:1px solid #dce7e3;border-radius:8px;padding:12px;min-height:50px}@media print{body{padding:0}.no-print{display:none}}</style></head><body><h1>🍒 Cherry Checklist</h1><p>Relatório completo da inspeção</p><h2>${escapeHtml(checklist.title)}</h2><div class="summary"><div><small>Resultado ponderado</small><br><b>${score}%</b></div><div><small>Progresso</small><br><b>${answered}/${checklist.questions.length}</b></div><div><small>Peso total</small><br><b>${totalWeight}%</b></div></div><h2>Identificação</h2><div class="meta"><div><b>Empresa:</b> ${escapeHtml(meta.company || "—")}</div><div><b>Responsável:</b> ${escapeHtml(meta.responsible || "—")}</div><div><b>Local:</b> ${escapeHtml(meta.location || "—")}</div><div><b>Data:</b> ${escapeHtml(meta.date || "—")}</div></div><h2>Avaliações</h2><table><thead><tr><th>#</th><th>Pergunta</th><th>Peso</th><th>Avaliação</th><th>Pontos</th></tr></thead><tbody>${rows}</tbody></table><h2>Observações</h2><div class="notes">${escapeHtml(meta.notes || "Nenhuma observação registrada.")}</div><p>Gerado pelo Cherry Checklist em ${new Date().toLocaleString("pt-BR")}.</p></body></html>`;
  };
  const saveCompleteReport = async () => {
    if (!isValid) { toast.error("Ajuste os pesos para 100% antes de salvar o relatório."); return; }
    const content = reportHtml(); const filename = reportFileName();
    try {
      const picker = (window as any).showSaveFilePicker;
      if (picker) { const handle = await picker({ suggestedName: filename, types: [{ description: "Relatório Cherry Checklist", accept: { "text/html": [".html"] } }] }); const writable = await handle.createWritable(); await writable.write(content); await writable.close(); toast.success("Relatório salvo na pasta escolhida."); return; }
    } catch (error) { if ((error as Error).name === "AbortError") return; }
    const url = URL.createObjectURL(new Blob([content], { type: "text/html;charset=utf-8" })); const link = document.createElement("a"); link.href = url; link.download = filename; link.click(); URL.revokeObjectURL(url); toast.success("Relatório completo baixado.");
  };
  const shareCompleteReport = async () => {
    if (!isValid) { toast.error("Ajuste os pesos para 100% antes de compartilhar."); return; }
    const file = new File([reportHtml()], reportFileName(), { type: "text/html" });
    try { if ((navigator as any).share && (!(navigator as any).canShare || (navigator as any).canShare({ files: [file] }))) { await (navigator as any).share({ title: checklist.title, text: `Relatório completo: ${score}%`, files: [file] }); return; } } catch { return; }
    await navigator.clipboard?.writeText(`Relatório: ${checklist.title} — resultado ${score}%\n${window.location.href}`); toast.success("Resumo copiado. Use Salvar relatório para enviar o arquivo completo.");
  };

  const printPdf = () => {
    if (!isValid) {
      toast.error("Ajuste os pesos para 100% antes de exportar o PDF.");
      return;
    }
    window.print();
  };

  const share = async () => {
    const text = `${checklist.title} — resultado ${score}% (${answered}/${checklist.questions.length} respondidas)`;
    if (navigator.share) {
      await navigator.share({ title: checklist.title, text });
    } else {
      await navigator.clipboard?.writeText(`${text}\n${window.location.href}`);
      toast.success("Resumo copiado para compartilhar.");
    }
  };

  const shareAppLink = async () => {
    const link = window.location.href;
    if (navigator.share) {
      await navigator.share({ title: "Cherry Checklist", text: "Abra o Cherry Checklist neste celular:", url: link });
    } else {
      await navigator.clipboard?.writeText(link);
      toast.success("Link do app copiado. Agora é só enviar.");
    }
  };

  if (!authReady) return <div className="auth-loading">🍒</div>;
  if (!unlocked) return <div className="auth-screen"><div className="auth-card"><div className="auth-logo">🍒</div><span className="section-kicker">CHERRY CHECKLIST</span><h1>{authMode === "create" ? "Crie seu acesso" : "Bem-vindo de volta"}</h1><p>{authMode === "create" ? "Cadastre um telefone e senha para proteger os dados salvos neste celular." : `Desbloqueie o checklist local${localAccount?.phone ? ` · ${localAccount.phone}` : ""}.`}</p><form onSubmit={submitAuth}><label className="field-label">Telefone<input type="tel" placeholder="(00) 00000-0000" value={authPhone} onChange={(e) => setAuthPhone(e.target.value)} disabled={authMode === "unlock"} /></label><label className="field-label">Senha<input type="password" placeholder="Mínimo de 6 caracteres" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} autoFocus /></label>{authMode === "create" && <label className="field-label">Confirme a senha<input type="password" placeholder="Digite novamente" value={authPasswordConfirm} onChange={(e) => setAuthPasswordConfirm(e.target.value)} /></label>}{authError && <div className="auth-error">{authError}</div>}<button className="full-button" type="submit">{authMode === "create" ? "Criar acesso e começar" : "Desbloquear app"}</button></form><small>Sem SMS, sem e-mail e sem nuvem. Seus dados ficam somente neste aparelho.</small>{authMode === "unlock" && <button className="auth-switch" onClick={() => { setAuthMode("create"); setAuthPassword(""); setAuthError(""); }}>Usar outro acesso neste aparelho</button>}</div></div>;

  return (
    <div className="app-shell">
      <header className="topbar print-hidden">
        <div className="brand-lockup">
          <div className="brand-mark"><span>🍒</span></div>
          <div><strong>Cherry</strong><span>Checklist</span></div>
        </div>
        <div className="top-actions">
          <button className="icon-button" aria-label="Ajuda" onClick={() => setShowHelp(true)}><CircleHelp size={20} /></button>
          <button className="icon-button menu-toggle" aria-label="Menu" onClick={() => setShowMenu(!showMenu)}><Menu size={20} /></button>
          <button className="install-button" onClick={installApp}><Download size={15} /> Instalar app</button><button className="primary-button compact" onClick={save}><Save size={17} /> Salvar</button>
        </div>
      </header>

      <main className="main-content">
        <div className="page-heading">
          <div>
            <div className="eyebrow"><span className="live-dot" /> CHECKLIST ATIVO <span className="code-pill">{checklist.code}</span></div>
            <h1>{checklist.title}</h1>
            <p className="muted">Monte, responda e compartilhe inspeções com rastreabilidade.</p>
          </div>
          <div className="heading-actions print-hidden">
            <button className="secondary-button" onClick={() => setActiveTab(activeTab === "checklist" ? "estrutura" : "checklist")}><Settings2 size={17} /> {activeTab === "checklist" ? "Editar estrutura" : "Responder checklist"}</button>
          </div>
        </div>

        <section className="score-grid">
          <div className="score-card score-primary">
            <div className="score-card-top"><span>Resultado ponderado</span><Gauge size={19} /></div>
            <div className="score-value">{score}<small>%</small></div>
            <div className="score-caption">de aproveitamento até agora</div>
            <div className="meter"><span style={{ width: `${score}%` }} /></div>
          </div>
          <div className="score-card">
            <div className="score-card-top"><span>Progresso</span><Check size={19} /></div>
            <div className="score-value dark">{progress}<small>%</small></div>
            <div className="score-caption">{answered} de {checklist.questions.length} perguntas respondidas</div>
            <div className="meter pale"><span style={{ width: `${progress}%` }} /></div>
          </div>
          <div className={`score-card ${isValid ? "weight-ok" : "weight-alert"}`}>
            <div className="score-card-top"><span>Peso configurado</span><Info size={19} /></div>
            <div className="score-value dark">{formatPercent(totalWeight)}<small>%</small></div>
            <div className="score-caption">{isValid ? "Distribuição pronta para fechar" : "Ajuste para fechar em 100%"}</div>
            <div className="weight-status">{isValid ? <><Check size={15} /> OK, total exato</> : <><Info size={15} /> Faltam {formatPercent(Math.max(0, 100 - totalWeight))}%</>}</div>
          </div>
        </section>

        <div className="content-grid">
          <section className="checklist-panel">
            <div className="panel-tabs print-hidden">
              <button className={activeTab === "checklist" ? "active" : ""} onClick={() => setActiveTab("checklist")}>Responder checklist</button>
              <button className={activeTab === "estrutura" ? "active" : ""} onClick={() => setActiveTab("estrutura")}>Editar estrutura</button>
            </div>

            {activeTab === "checklist" ? (
              <div className="question-list">
                <div className="section-heading"><div><span className="section-kicker">ETAPA 01</span><h2>Verificação dos itens</h2></div><span className="count-badge">{answered}/{checklist.questions.length}</span></div>
                <p className="section-description">Marque <b>Conforme</b>, <b>Não conforme</b>, <b>Não se aplica</b> ou <b>Não observado</b>.</p>
                {checklist.questions.map((q, index) => (
                  <React.Fragment key={q.id}>{(index === 0 || q.section !== checklist.questions[index - 1]?.section) && <div className="topic-heading"><span className="topic-number">{q.section?.split(".")[0]}</span><div><span className="section-kicker">TÓPICO</span><h3>{q.section}</h3></div></div>}
                  <article className={`question-card ${q.status ? "answered" : ""}`}>
                    <div className="question-index">{String(index + 1).padStart(2, "0")}</div>
                    <div className="question-body"><p>{q.text}</p><span className="weight-label">Peso {q.weight}%</span></div>
                    <div className="answer-actions">
                      <button className={`answer-button conform ${q.status === "conforme" ? "selected" : ""}`} onClick={() => updateQuestion(q.id, { status: q.status === "conforme" ? null : "conforme" })}><Check size={18} /> <span>Conforme</span></button>
                      <button className={`answer-button nonconform ${q.status === "nao-conforme" ? "selected" : ""}`} onClick={() => updateQuestion(q.id, { status: q.status === "nao-conforme" ? null : "nao-conforme" })}><X size={18} /> <span>Não conforme</span></button>
                      <button className={`answer-button notapplicable ${q.status === "nao-se-aplica" ? "selected" : ""}`} onClick={() => updateQuestion(q.id, { status: q.status === "nao-se-aplica" ? null : "nao-se-aplica" })}><Minus size={18} /> <span>Não se aplica</span></button>
                      <button className={`answer-button notobserved ${q.status === "nao-observado" ? "selected" : ""}`} onClick={() => updateQuestion(q.id, { status: q.status === "nao-observado" ? null : "nao-observado" })}><Info size={18} /> <span>Não observado</span></button>
                    </div>
                    {q.status && <div className="question-attachments"><div className="attachment-heading"><span>Fotos da evidência</span><small>{q.photos?.length || 0}/4</small></div><div className="photo-strip">{(q.photos || []).map((photo, photoIndex) => <div className={`photo-thumb ${photo.isPrimary ? "primary-photo" : ""}`} key={`${q.id}-${photoIndex}`}><img src={photo.src} alt={`Evidência ${photoIndex + 1}`} /><button type="button" onClick={() => removePhoto(q.id, photoIndex)} aria-label="Remover foto">×</button><span>{photo.isPrimary ? "Principal" : ""}</span></div>)}<label className="add-photo"><Upload size={16} /><span>Tirar foto</span><input type="file" accept="image/*" capture="environment" hidden onChange={(event) => { addPhotos(q.id, event.target.files); event.currentTarget.value = ""; }} /></label><label className="add-photo"><FileUp size={16} /><span>Anexar foto</span><input type="file" accept="image/*" multiple hidden onChange={(event) => { addPhotos(q.id, event.target.files); event.currentTarget.value = ""; }} /></label></div>{(q.photos || []).map((photo, photoIndex) => <div className="photo-details" key={`${q.id}-details-${photoIndex}`}><input placeholder="Comentário da foto" value={photo.comment} onChange={(event) => updatePhoto(q.id, photoIndex, { comment: event.target.value })} /><button type="button" className={photo.isPrimary ? "primary-photo-button active" : "primary-photo-button"} onClick={() => updatePhoto(q.id, photoIndex, { isPrimary: !photo.isPrimary })}>{photo.isPrimary ? "Foto principal" : "Definir como principal"}</button><small>{new Date(photo.capturedAt).toLocaleString("pt-BR")}</small></div>)}<small className="attachment-note">Até 4 fotos por pergunta. Escolha tirar uma nova foto ou anexar uma da galeria.</small></div>}
                  </article></React.Fragment>
                ))}
                {!checklist.questions.length && <div className="empty-state">Nenhuma pergunta cadastrada. Vá em editar estrutura para começar.</div>}
              </div>
            ) : (
              <div className="structure-editor">
                <div className="section-heading"><div><span className="section-kicker">ETAPA 02</span><h2>Estrutura do checklist</h2></div><span className={`count-badge ${isValid ? "success" : "warning"}`}>{formatPercent(totalWeight)}% / 100%</span></div>
                <p className="section-description">Edite as perguntas e distribua os pesos. A soma precisa fechar exatamente em 100%.</p>
                <label className="field-label">Nome do checklist<input value={checklist.title} onChange={(e) => setChecklist({ ...checklist, title: e.target.value })} /></label>
                <div className="meta-section">
                  <div className="meta-section-heading"><div><span className="section-kicker">DADOS DA INSPEÇÃO</span><h3>Identificação e rastreabilidade</h3></div><span className="mini-note">Opcional</span></div>
                  <div className="meta-grid">
                    <label className="field-label">Responsável<input placeholder="Nome de quem preenche" value={checklist.meta?.responsible || ""} onChange={(e) => setChecklist({ ...checklist, meta: { ...starter.meta!, ...checklist.meta, responsible: e.target.value } })} /></label>
                    <label className="field-label">Empresa<input placeholder="Nome da empresa" value={checklist.meta?.company || ""} onChange={(e) => setChecklist({ ...checklist, meta: { ...starter.meta!, ...checklist.meta, company: e.target.value } })} /></label>
                    <label className="field-label">Local<input placeholder="Unidade ou endereço" value={checklist.meta?.location || ""} onChange={(e) => setChecklist({ ...checklist, meta: { ...starter.meta!, ...checklist.meta, location: e.target.value } })} /></label>
                    <label className="field-label">Data<input type="date" value={checklist.meta?.date || ""} onChange={(e) => setChecklist({ ...checklist, meta: { ...starter.meta!, ...checklist.meta, date: e.target.value } })} /></label>
                    <div className="field-label"><span>Assinatura / conferência</span><SignaturePad value={checklist.meta?.signature || ""} onChange={(signature) => setChecklist({ ...checklist, meta: { ...starter.meta!, ...checklist.meta, signature } })} /></div>
                    <label className="field-label">Observações<textarea placeholder="Anote achados, ações ou recomendações" value={checklist.meta?.notes || ""} onChange={(e) => setChecklist({ ...checklist, meta: { ...starter.meta!, ...checklist.meta, notes: e.target.value } })} /></label>
                  </div>
                </div>
                <div className="editor-list">
                  {checklist.questions.map((q, index) => <React.Fragment key={q.id}>{(index === 0 || q.section !== checklist.questions[index - 1]?.section) && <div className="topic-heading editor-topic"><span className="topic-number">{q.section?.split(".")[0]}</span><div><span className="section-kicker">TÓPICO</span><h3>{q.section}</h3></div></div>}<div className="editor-row"><GripVertical size={18} className="drag-icon" /><span className="editor-number">{index + 1}</span><input className="question-input" value={q.text} onChange={(e) => updateQuestion(q.id, { text: e.target.value })} /><label className="weight-input"><input type="number" min="0" max="100" value={q.weight} onChange={(e) => updateQuestion(q.id, { weight: normalize(e.target.value) })} /><span>%</span></label><button className="delete-button" onClick={() => removeQuestion(q.id)} aria-label="Excluir pergunta"><Trash2 size={17} /></button></div></React.Fragment>)}
                </div>
                <button className="add-question" onClick={addQuestion}><Plus size={18} /> Adicionar pergunta</button>
              </div>
            )}
          </section>

          <aside className="side-panel print-hidden">
            <div className="side-card export-card"><div className="side-icon cherry"><Download size={19} /></div><div><h3>Relatório completo</h3><p>Salve ou envie o checklist preenchido com avaliações e percentuais.</p></div><button className="full-button" onClick={printPdf}><FileDown size={17} /> Gerar PDF</button><button className="outline-button" onClick={saveCompleteReport}><Save size={16} /> Salvar na pasta</button><button className="outline-button" onClick={shareCompleteReport}><Share2 size={16} /> Enviar arquivo completo</button></div>
            <div className="side-card"><div className="side-icon green"><Share2 size={19} /></div><div><h3>Compartilhar resultado</h3><p>Envie pelo WhatsApp, e-mail ou outro aplicativo.</p></div><button className="outline-button" onClick={share}><Share2 size={17} /> Compartilhar</button><a className="text-link" href={`https://wa.me/?text=${encodeURIComponent(`${checklist.title} — resultado ${score}%`)}`} target="_blank" rel="noreferrer">Abrir no WhatsApp <ChevronRight size={15} /></a><a className="text-link" href={`mailto:?subject=${encodeURIComponent(checklist.title)}&body=${encodeURIComponent(`Resultado: ${score}%`)}`}><Mail size={15} /> Enviar por e-mail <ChevronRight size={15} /></a></div>
            <div className="side-card import-card"><div className="side-icon blue"><Upload size={19} /></div><div><h3>Arquivos de perguntas</h3><p>Digite manualmente ou importe perguntas prontas em DOC/DOCX, TXT, PDF, Excel, CSV ou JSON.</p></div><div className="two-buttons"><button className="outline-button" onClick={loadModelChecklist}><Check size={16} /> Usar modelo Boas Práticas</button><button className="outline-button" onClick={exportQuestionsCsv}><FileUp size={16} /> Excel / CSV</button><button className="outline-button" onClick={exportQuestionsWord}><FileUp size={16} /> Word</button><label className="outline-button"><FileDown size={16} /> Importar<input key={checklist.questions.length} type="file" hidden onChange={(e) => importQuestionsFile(e.target.files?.[0])} /></label></div><div className="file-actions"><button className="text-link" onClick={downloadQuestionTemplate}>Baixar modelo Excel/CSV <ChevronRight size={15} /></button><button className="text-link" onClick={downloadWordTemplate}>Baixar modelo Word: Boas Práticas <ChevronRight size={15} /></button><button className="text-link" onClick={exportJson}>Exportar estrutura completa <ChevronRight size={15} /></button></div></div><div className="side-card app-link-card"><div className="side-icon blue"><Share2 size={19} /></div><div><h3>Enviar para outro celular</h3><p>Compartilhe o link para abrir e instalar o Cherry Checklist em outro aparelho.</p></div><button className="outline-button" onClick={shareAppLink}><Share2 size={17} /> Compartilhar link do app</button><span className="local-note">O link instala o app; os dados permanecem neste celular.</span></div>
            <div className="side-card history-card"><div className="side-icon purple"><Gauge size={19} /></div><div><h3>Histórico e relatórios</h3><p>Filtre por período, empresa ou responsável.</p></div><div className="history-filters"><input className="profile-input" placeholder="Empresa" value={filterCompany} onChange={(e) => setFilterCompany(e.target.value)} /><input className="profile-input" placeholder="Responsável" value={filterResponsible} onChange={(e) => setFilterResponsible(e.target.value)} /><div className="filter-dates"><input className="profile-input" type="date" value={filterFrom} onChange={(e) => setFilterFrom(e.target.value)} /><input className="profile-input" type="date" value={filterTo} onChange={(e) => setFilterTo(e.target.value)} /></div></div>{filteredHistory.length ? <div className="history-list">{filteredHistory.slice(0, 4).map((item) => <div className="history-item" key={item.id}><div><b>{item.title}</b><span>{new Date(item.date).toLocaleDateString("pt-BR")} · {item.company || "Sem empresa"}</span></div><strong>{item.score}%</strong></div>)}</div> : <div className="empty-history">Nenhum registro encontrado com esses filtros.</div>}<div className="chart-bars" aria-label="Evolução das inspeções">{filteredHistory.slice(0, 6).reverse().map((item) => <span key={item.id} style={{ height: `${Math.max(12, item.score)}%` }} title={`${item.score}%`} />)}</div><div className="two-buttons report-buttons"><button className="outline-button" onClick={exportHistoryCsv}><FileUp size={15} /> Excel / CSV</button><button className="outline-button" onClick={exportHistoryPdf}><FileDown size={15} /> PDF</button></div></div>
            <div className="side-card profile-card"><div className="side-icon orange"><Settings2 size={19} /></div><div><h3>Perfil local</h3><p>Identificação opcional salva somente neste celular.</p></div><input className="profile-input" placeholder="Seu nome" value={profile.name} onChange={(e) => setProfile({ ...profile, name: e.target.value })} /><input className="profile-input" placeholder="Seu e-mail" type="email" value={profile.email} onChange={(e) => setProfile({ ...profile, email: e.target.value })} /><span className="local-badge"><Check size={14} /> Sem login e sem nuvem</span><button className="outline-button" onClick={() => setShowPasswordChange(!showPasswordChange)}><Settings2 size={15} /> Alterar senha</button>{showPasswordChange && <div className="password-change"><input className="profile-input" type="password" placeholder="Nova senha" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} /><input className="profile-input" type="password" placeholder="Confirme a nova senha" value={newPasswordConfirm} onChange={(e) => setNewPasswordConfirm(e.target.value)} /><button className="full-button" onClick={changePassword}>Salvar nova senha</button></div>}</div>
          </aside>
        </div>
      </main>

      <footer className="footer print-hidden"><span>🍒 Cherry Checklist</span><span>Seus dados ficam salvos neste dispositivo.</span><button onClick={() => setShowHelp(true)}>Como funciona?</button></footer>
      {showMenu && <div className="mobile-menu print-hidden"><button onClick={() => { setActiveTab("estrutura"); setShowMenu(false); }}><Settings2 size={17} /> Configurações e perguntas</button><button onClick={exportJson}><FileUp size={17} /> Exportar estrutura</button><button onClick={() => setShowHelp(true)}><CircleHelp size={17} /> Ajuda</button></div>}
      {showHelp && <div className="modal-backdrop print-hidden" onClick={() => setShowHelp(false)}><div className="help-modal" onClick={(e) => e.stopPropagation()}><div className="modal-cherry">🍒</div><h2>Como usar o Cherry Checklist</h2><p>Crie suas perguntas, distribua pesos até totalizar 100%, responda os itens e baixe o resultado em PDF. O checklist fica salvo automaticamente no navegador deste dispositivo.</p><button className="full-button" onClick={() => setShowHelp(false)}>Entendi</button></div></div>}
    </div>
  );
}
