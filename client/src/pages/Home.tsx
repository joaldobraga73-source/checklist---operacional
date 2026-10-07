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
  RotateCcw,
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
  { id: "odt-1", text: "1.1 Os colaboradores que apresentam feridas, lesões, gripe ou outra situação de risco, são afastados da manipulação de alimentos e recebem orientações de como proceder?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-2", text: "1.2 Os colaboradores que apresentam feridas, lesões, gripe ou outra situação de risco, recebem luva e máscara para proteção?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-3", text: "1.3 Os exames periódicos, ASO,dos manipuladores estão atualizados?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-4", text: "1.4 Os manipuladores higienizam as mãos com frequência, sempre que necessário?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-5", text: "1.5 Os colaboradores estão uniformizados, com uniformes limpos e em bom estado de conservação?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-6", text: "1.6 Todos que acessam a área de produção, manipuladores e visitantes usam toucas ou boné para proteção dos cabelos?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-7", text: "1.7 Os manipuladores estão sem adornos (anéis, brincos, relógios, pulseiras, correntes ou similares)?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-8", text: "1.8 Todos os colaboradores usam EPIs, sapato de segurança fechado e antiderrapante, avental, luva anticorte limpos e em bom estado de conservação e limpeza?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-9", text: "1.9 Manipuladores mantem boa higiene pessoal, cabelo e barba aparadas, unhas curtas e sem esmalte?", weight: 1.32, status: null, section: "1. SAÚDE E HIGIENE DOS MANIPULADORES" },
  { id: "odt-10", text: "2.1 O piso da Unidade está limpo?", weight: 1.32, status: null, section: "2. PISO, PAREDES, TETO, PORTAS E JANELAS" },
  { id: "odt-11", text: "2.2 Os ralos da cozinha estão limpos, desobstruídos e funcionando?", weight: 1.32, status: null, section: "2. PISO, PAREDES, TETO, PORTAS E JANELAS" },
  { id: "odt-12", text: "2.3 As paredes da Unidade são de material liso, impermeável, lavável, em bom estado de conservação e limpeza?", weight: 1.32, status: null, section: "2. PISO, PAREDES, TETO, PORTAS E JANELAS" },
  { id: "odt-13", text: "2.4 O teto e forro da Unidade são mantidos íntegros, conservados e limpos, livre de bolores e teias de aranha?", weight: 1.32, status: null, section: "2. PISO, PAREDES, TETO, PORTAS E JANELAS" },
  { id: "odt-14", text: "2.5 As portas, janelas e telas da Unidade estão limpas?", weight: 1.32, status: null, section: "2. PISO, PAREDES, TETO, PORTAS E JANELAS" },
  { id: "odt-15", text: "2.6 Dentro da área de produção as lâmpadas e protetores estão limpos?", weight: 1.32, status: null, section: "2. PISO, PAREDES, TETO, PORTAS E JANELAS" },
  { id: "odt-16", text: "3.1 Os vasos sanitários de colaboradores estão tampados e limpos?", weight: 1.32, status: null, section: "3. SANITÁRIOS E VESTIÁRIOS" },
  { id: "odt-17", text: "3.2 Nos sanitários: os facilitadores (toalha de papel descartável não reciclável, sabonete líquido inodoro, antisséptico) estão abastecidos?", weight: 1.32, status: null, section: "3. SANITÁRIOS E VESTIÁRIOS" },
  { id: "odt-18", text: "3.3 As lixeiras dos sanitários estão limpas, com tampa acionada por pedale com capacidade adequada, sem excesso de lixo?", weight: 1.32, status: null, section: "3. SANITÁRIOS E VESTIÁRIOS" },
  { id: "odt-19", text: "3.4 Nos sanitários há presença de avisoscom a instruçãopara higiene das mãos?", weight: 1.32, status: null, section: "3. SANITÁRIOS E VESTIÁRIOS" },
  { id: "odt-20", text: "3.5 Os armários de colaboradores estão fechados e os pertences pessoais guardados, em bom estado de conservação e limpeza?", weight: 1.32, status: null, section: "3. SANITÁRIOS E VESTIÁRIOS" },
  { id: "odt-21", text: "3.6 Os banheiros estão limpos e são higienizados somente ao final do expediente ou por equipe exclusiva de limpeza, com escala de limpeza preenchida e atualizada?", weight: 1.32, status: null, section: "3. SANITÁRIOS E VESTIÁRIOS" },
  { id: "odt-22", text: "4.1 Os cestos de lixo estão limpos, tampados, com acionamento por pedal,em número compatível com a produção, sem excessode lixo acumulado?", weight: 1.32, status: null, section: "4. MANEJO RESÍDUOS E CONTROLE DE PRAGAS" },
  { id: "odt-23", text: "4.3 As caixas de gordura são limpas a cada 3 meses e possuem registro?", weight: 1.32, status: null, section: "4. MANEJO RESÍDUOS E CONTROLE DE PRAGAS" },
  { id: "odt-24", text: "4.4 Osresíduos orgânicos são separados dosrecicláveis(coleta seletiva)?", weight: 1.32, status: null, section: "4. MANEJO RESÍDUOS E CONTROLE DE PRAGAS" },
  { id: "odt-25", text: "4.5 Os arredores da Unidade apresentam ausência de animais domésticos?", weight: 1.32, status: null, section: "4. MANEJO RESÍDUOS E CONTROLE DE PRAGAS" },
  { id: "odt-26", text: "4.6 Os dispositivos, caixas de iscas estão devidamente posicionados e em bom estado de conservação?", weight: 1.32, status: null, section: "4. MANEJO RESÍDUOS E CONTROLE DE PRAGAS" },
  { id: "odt-27", text: "4.7 O depósito de lixo externo é isolado, tampado ou tratado de forma a evitar contaminação?", weight: 1.32, status: null, section: "4. MANEJO RESÍDUOS E CONTROLE DE PRAGAS" },
  { id: "odt-28", text: "5.1 É realizada a higienização completa dos hortifrútis, lavagem, desinfecção com sanitizante, cloro e enxague em pia ou bancada limpa e exclusiva para manipulação de alimentos?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-29", text: "5.2 Na área de produção as bancadas e pias são lisas, de material impermeável, lavável, estão limpas organizadas e conservadas?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-30", text: "5.3 Na área de produção os equipamentos estão limpos organizados e conservados?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-31", text: "5.4 A área de produção é livre de materiais em desuso, papelãoe madeira?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-32", text: "5.5 Existe termômetro com calibração vigente para monitorar a temperatura do alimento?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-33", text: "5.6 Existe planilha diária de controle de temperaturapreenchida?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-34", text: "5.7 Existe pia exclusiva e desobstruída, dotada de papel, sabonete líquido, álcool 70% e lixeira com tampa e acionamento por pedal para higienização das mãos dentro da produção?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-35", text: "5.8 Os manipuladores higienizam as mãos e fazem troca de luvas frequentemente e sempre que necessário?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-36", text: "5.9 Na cozinha a disposição de armazenamento segue PEPS (primeiro que entra, primeiro que sai)?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-37", text: "5.10 Os produtos em embalagens danificadas, estragados ou vencidos são armazenados em local separado e identificado para descarte?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-38", text: "5.11 Os alimentos armazenados transferidos de embalagem, abertosoufracionados, estão tampados, protegidos e identificados com prazo de validade?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-39", text: "5.12 Todo produto armazenado está dentro do prazo de validade?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-40", text: "5.13 A área de produção está livre de vestígios de moscas, insetos, pragas e roedores?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-41", text: "5.14 Na área de produção não há degustação de alimentos?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-42", text: "5.15 A área de manipulação está livre de utilização de celular e carregadores de celularnas bancadas?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-43", text: "5.16 Os produtos estocados na cozinha ou em manipulação estão sob paletes (de no mínimo15cm), sem contato direto com o chão?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-44", text: "5.17 As geladeiras e freezes da cozinha estão organizados, limpos, sem excesso de gelo, sem presença de materiais pessoais e os alimentos separados por categoria, evitando contaminação cruzada?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-45", text: "5.18 Possui ventiladores ligados na cozinha, com corrente de ar de forma a não incidir sobre os alimentos nas áreas de preparo durante a manipulação dos alimentos?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-46", text: "5.19 Os colaboradores usam luvas de limpeza para retirar o lixo e higienizam a mão ao retornar para cozinha?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-47", text: "5.20 Os óleos e gorduras das frituras são aquecidos no máximo até 180ºC, sendodescartados quandoobservado o ponto de saturação, mudança de cor, presença de espuma e fumaça?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-48", text: "5.21 O óleo das frituras quando desprezado obedece a regras de sustentabilidade, não sendo descartados no meio ambiente, são armazenados em recipientes com tampas, em local apropriado enquanto aguardam o recolhimento?", weight: 1.32, status: null, section: "5. ÁREA DE PRE-PREPARO E PREPARO DE ALIMENTOS" },
  { id: "odt-49", text: "6.1 Na área de produção possui borrifadores de solução clorada, água sanitáriadisponívelpara desinfecção de ambiente, alimentos e equipamentos?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-50", text: "6.2 Existe registro (POP) de limpeza preenchido corretamente?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-51", text: "6.3 A Unidade está livre de produtos de limpeza, spray, aerossóis, odorizantes ou desodorantes?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-52", text: "6.4 A área de produção está livre de materiais de limpeza, vassoura, rodos, panos e similares?Os mesmos estão armazenados em local adequado fora da produção?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-53", text: "6.5 Os utensíliose equipamentosutilizados estão limpos e são devidamente higienizadosdiariamente, sem etiquetas antigase resíduos de alimentos?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-54", text: "6.6 Os armários de massas são varridos diariamente ao final do expediente e retirados o excesso de resíduos, são lavados semanalmente?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-55", text: "6.7 As pás de madeira são raspadas diariamente, estão limpas sem presença de resíduos e em bom estado de conservação?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-56", text: "6.8 A coifa está limpa, sem gotejamento e sujidadesaparentes?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-57", text: "6.9 A pia de higienização de utensílios está sem acúmulo excessivo, o volume de louça suja está compatível com o espaço?", weight: 1.32, status: null, section: "6. HIGIENIZAÇÃO DE INSTALAÇÕES, EQUIPAMENTOS, MÓVEIS E UTENSÍLIOS" },
  { id: "odt-58", text: "7.1 No recebimento é conferida nota fiscal,temperatura, quantidade, qualidade, datas de validade, condições das embalagens, avaliação sensorial (cor, cheiro e aparência) dos produtos?", weight: 1.32, status: null, section: "7. RECEBIMENTO DE MERCADORIA" },
  { id: "odt-59", text: "7.2 No recebimento existe ordem de prioridade para armazenar os produtos? Perecíveis primeiro, com tempo de exposição à temperatura ambiente inferior a 30 minutos?", weight: 1.32, status: null, section: "7. RECEBIMENTO DE MERCADORIA" },
  { id: "odt-60", text: "7.3 No recebimento existecuidadopara as mercadorias ficarem sob paletes, sem contato direto com o piso?", weight: 1.32, status: null, section: "7. RECEBIMENTO DE MERCADORIA" },
  { id: "odt-61", text: "8.1 No estoque os alimentos são armazenados de acordo com suas especificações, separados por tipo:perecível e não perecível, ou seja,congelado, refrigerado e seco?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-62", text: "8.2 No estoque os produtos estão dispostos como PEPS (primeiro que entra, primeiro que sai)?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-63", text: "8.3 As áreas de armazenamentos e estoque estão em bom estado de conservação, prateleiras limpas, livre de materiais em desuso e excesso de caixas de papelão?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-64", text: "8.4 A câmara fria ou refrigerador está organizado? Alimentos prontos pra consumo, cozidos separados dos crus, carnese derivadosseparados doshortifrutis?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-65", text: "8.6 No estoque os produtos de limpeza e descartáveis estão separados dos alimentos?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-66", text: "8.9 No estoque e câmaras os produtos em embalagens danificadas, estragados ou vencidos são armazenados em local separado e identificado para descarte?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-67", text: "8.10 No estoque e câmaras os alimentos armazenados transferidos de embalagem, abertos e fracionados estão rotulados, identificados, datados com prazo de validade vigente e protegidos, sem alimentos vencidos?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-68", text: "8.11 A área de estoque e câmaras estão livres devestígiosde pragas, moscas, insetos e roedores?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-69", text: "8.12 A área de armazenamento e câmaras está livre de degustação de alimentos?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-70", text: "8.13 No estoque e câmaras, todos os produtos estão sob paletes (15cm do piso), sem contato direto com o chão, a 5cm da parede e 45cm do teto?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-71", text: "8.14 O descongelamento é feito sob refrigeração, quando possível, protegido, identificados com etiquetas de descongelamento e prazo de até 3 dias para descongelamento?", weight: 1.32, status: null, section: "8. ESTOQUE E ARMAZENAMENTO" },
  { id: "odt-72", text: "9.1 No transporte de mercadorias e alimentos prontos o veículo apresenta proteção de carga?", weight: 1.32, status: null, section: "9. TRANSPORTE DE ALIMENTOS" },
  { id: "odt-73", text: "9.2 O veículo de transporte é revestido de material lavável, refrigerado, em bom estado de conservação e limpeza?", weight: 1.32, status: null, section: "9. TRANSPORTE DE ALIMENTOS" },
  { id: "odt-74", text: "10.1 No refeitório de colaboradores a área está limpa e organizada, não havendo desperdício de alimentos?", weight: 1.32, status: null, section: "10. ÁREA DE REFEITÓRIO DE COLABORADORES" },
  { id: "odt-75", text: "10.2 O bebedouro de colaboradores está em bom estado de conservação, limpeza e os filtros são trocados há cada seis meses com registro de troca?", weight: 1.32, status: null, section: "10. ÁREA DE REFEITÓRIO DE COLABORADORES" },
  { id: "odt-76", text: "10.3 As lixeiras do bebedouro e área do refeitório de colaboradores estão tampadas com capacidade adequada, sem excesso de lixo?", weight: 1, status: null, section: "10. ÁREA DE REFEITÓRIO DE COLABORADORES" },
].map((question) => ({ ...question, section: question.section || sectionForQuestion(question.text) }));

const starter: Checklist = {
  title: "Checklist de Boas Práticas",
  code: "CK-BOAS-PRATICAS",
  questions: modelQuestions,
  meta: { responsible: "", company: "", location: "", date: new Date().toISOString().slice(0, 10), signature: "", notes: "" },
};

const uid = () => `q-${Date.now()}-${Math.random().toString(16).slice(2)}`;
const normalize = (value: string) => Number(value.replace(",", ".")) || 0;
const normalizeChecklist = (value: any): Checklist => ({ ...value, questions: (value.questions || []).map((question: any) => { const text = String(question.text || "").replace(/^1\.5 É realizada higienização das embalagens/, "14.2 É realizada higienização das embalagens").replace(/^14\.29 Na área de produção possui borrifadores/, "15.1 Na área de produção possui borrifadores"); return { ...question, text, section: question.section || sectionForQuestion(text), photos: (question.photos || []).map((photo: any) => typeof photo === "string" ? { src: photo, comment: "", isPrimary: false, capturedAt: new Date().toISOString() } : photo) }; }) });
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
  const applicableWeight = useMemo(() => checklist.questions.reduce((sum, q) => sum + (q.status !== "nao-se-aplica" && q.status !== "nao-observado" ? q.weight : 0), 0), [checklist.questions]);
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
    toast.success(`Modelo de Boas Práticas carregado com ${modelQuestions.length} perguntas.`);
  };

  const startNewReport = () => {
    if (!window.confirm("Começar um novo relatório? A avaliação atual permanecerá no histórico, mas as respostas e fotos desta tela serão limpas.")) return;
    setChecklist({ ...starter, questions: modelQuestions.map((question) => ({ ...question, id: uid() })), meta: { ...starter.meta!, date: new Date().toISOString().slice(0, 10) } });
    setActiveTab("checklist");
    toast.success("Novo relatório iniciado.");
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

  const reportFileName = (kind: "conformidades" | "nao-conformidades" = "nao-conformidades") => `${checklist.title.toLowerCase().replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "cherry-checklist"}-relatorio-${kind}.html`;
  const escapeHtml = (value: string) => value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[char] || char));
  const scoreTone = (value: number) => value >= 80 ? { label: "Verde", className: "green" } : value >= 60 ? { label: "Amarelo", className: "yellow" } : { label: "Vermelho", className: "red" };
  const reportHtml = (kind: "conformidades" | "nao-conformidades" = "nao-conformidades") => {
    const meta = checklist.meta || starter.meta!;
    const tone = scoreTone(score);
    const grade = formatPercent(score / 10);
    const topics = checklist.questions.reduce<Record<string, Question[]>>((groups, question) => {
      const topic = question.section || "Outros itens";
      groups[topic] = [...(groups[topic] || []), question];
      return groups;
    }, {});
    const topicRows = Object.entries(topics).map(([topic, questions]) => {
      const possible = questions.reduce((sum, question) => sum + (question.status !== "nao-se-aplica" && question.status !== "nao-observado" ? question.weight : 0), 0);
      const obtained = questions.reduce((sum, question) => sum + (question.status === "conforme" ? question.weight : 0), 0);
      const percentage = possible ? (obtained / possible) * 100 : null;
      return `<tr><td>${escapeHtml(topic)}</td><td>${formatPercent(possible)}</td><td>${percentage === null ? "—" : formatPercent(obtained)}</td><td>${percentage === null ? "Nota não calculada" : `${formatPercent(percentage)} %`}</td></tr>`;
    }).join("");
    const findings = checklist.questions.filter((question) => kind === "conformidades" ? question.status === "conforme" : question.status === "nao-conforme");
    const resultRows = findings.length ? findings.map((question) => {
      const photos = kind === "nao-conformidades" ? (question.photos || []).map((photo) => `<figure><img src="${photo.src}" alt="Evidência" /><figcaption>${photo.isPrimary ? "Foto principal · " : ""}${escapeHtml(photo.comment || "Sem comentário")}</figcaption></figure>`).join("") : "";
      const observation = (question.photos || []).map((photo) => photo.comment).filter(Boolean).join(" | ");
      return `<article class="finding"><div class="finding-head"><span class="finding-number">${escapeHtml(question.text.split(" ")[0] || "—")}</span><strong>${kind === "conformidades" ? "Conforme" : "Não conforme"}</strong></div><p class="finding-question">${escapeHtml(question.text)}</p>${kind === "nao-conformidades" && observation ? `<p class="observation"><b>Observação:</b> ${escapeHtml(observation)}</p>` : ""}${photos ? `<div class="photo-grid">${photos}</div>` : ""}</article>`;
    }).join("") : `<div class="empty-result">Nenhum item ${kind === "conformidades" ? "conforme" : "não conforme"} registrado.</div>`;
    const reportTitle = kind === "conformidades" ? "Relatório de conformidades" : "Relatório de não conformidades";
    const statusDescription = kind === "conformidades" ? "Itens avaliados como conformes" : "Itens avaliados como não conformes, com evidências fotográficas";
    return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>${escapeHtml(checklist.title)} — ${reportTitle}</title><style>
      @page{size:A4;margin:12mm 13mm 17mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#17252b;margin:0;border-top:4px solid #0c5b22;font-size:11px}header{padding:18px 0 10px;border-bottom:1px solid #9aa9a0}header h1{font-size:22px;line-height:1.05;margin:0;color:#1e3434;text-transform:uppercase;max-width:520px}header p{color:#9b4b4b;font-size:10px;margin:7px 0 0;text-transform:uppercase}h2{font-size:15px;text-align:center;color:#284849;margin:17px 0 8px;text-transform:uppercase;letter-spacing:.02em}h2:before,h2:after{content:"";display:inline-block;vertical-align:middle;width:25%;border-top:1px solid #84938b;margin:0 12px}.meta-grid{display:grid;grid-template-columns:1fr 1fr;gap:0 14px;margin-top:8px}.meta-grid div{display:grid;grid-template-columns:92px 1fr;min-height:22px;align-items:center}.meta-grid b{background:#075b22;color:#fff;padding:4px 6px;text-transform:uppercase;font-size:9px}.meta-grid span{padding:4px 7px;background:#f2f3f2}.score-box{display:flex;justify-content:space-between;gap:12px;align-items:stretch;margin-top:8px}.score-main{flex:1;padding:12px 16px;border-radius:4px;background:#e5f4e2;border-left:6px solid #2f8d30}.score-main.yellow{background:#fff7c7;border-left-color:#e0b900}.score-main.red{background:#ffe1df;border-left-color:#d52b2b}.score-main .label{font-size:10px;text-transform:uppercase;font-weight:bold}.score-main .value{font-size:27px;font-weight:800;margin-top:3px}.score-main .value span{font-size:15px}.score-main .tone{font-size:12px;font-weight:bold}.geo{margin-top:8px;border-top:1px solid #d0d7d3;padding-top:6px;font-size:9px}.geo b{display:inline-block;width:42px}.summary-table,.result-table{width:100%;border-collapse:collapse;margin-top:6px}.summary-table th,.summary-table td{padding:5px 6px;border-bottom:1px solid #cfd9d2;text-align:left}.summary-table th{background:#075b22;color:#fff;font-size:10px}.summary-table td:nth-child(n+2){text-align:right}.finding{border-top:1px dashed #77877e;padding:10px 0 12px;break-inside:avoid}.finding-head{display:flex;justify-content:space-between;align-items:center}.finding-number{font-weight:bold;color:#fff;background:#075b22;padding:4px 8px}.finding-head strong{color:#b52c2c;font-size:10px;text-transform:uppercase}.finding-question{font-size:11px;line-height:1.38;margin:5px 0;color:#27383a}.observation{font-size:10px;color:#9b4b4b;margin:5px 0}.photo-grid{display:flex;flex-wrap:wrap;gap:10px;margin-top:7px}.photo-grid figure{margin:0;width:calc(50% - 5px);border:1px solid #8e9992;padding:7px;text-align:center;break-inside:avoid}.photo-grid img{display:block;width:100%;max-width:100%;height:280px;object-fit:contain;background:#f7f8f7}.photo-grid figcaption{font-size:8px;color:#687873;margin-top:3px}.empty-result{padding:14px;background:#e5f4e2;color:#075b22;text-align:center;font-weight:bold}.notes{white-space:pre-wrap;border:1px solid #cfd9d2;padding:9px;min-height:35px}.footer{margin-top:18px;padding-top:7px;border-top:3px solid #075b22;display:flex;justify-content:space-between;color:#687873;font-size:8px}@media print{.no-print{display:none}}
    </style></head><body><header><h1>Lista de verificação do serviço de alimentação</h1><p>${statusDescription}</p></header><section class="meta-grid"><div><b>Data</b><span>${escapeHtml(meta.date || new Date().toLocaleDateString("pt-BR"))}</span></div><div><b>Nota</b><span>${grade} — ${tone.label}</span></div><div><b>Avaliador</b><span>${escapeHtml(meta.responsible || "—")}</span></div><div><b>Avaliado</b><span>${escapeHtml(meta.company || checklist.title)}</span></div><div><b>Local</b><span>${escapeHtml(meta.location || "—")}</span></div><div><b>Itens</b><span>${answered}/${checklist.questions.length} respondidos</span></div></section><div class="score-box"><div class="score-main ${tone.className}"><div class="label">Nota final</div><div class="value">${grade}<span>/10</span></div><div class="tone">Classificação: ${tone.label}</div></div><div class="score-main"><div class="label">Progresso da avaliação</div><div class="value">${progress}<span>%</span></div><div class="tone">${answered} de ${checklist.questions.length} perguntas</div></div></div><h2>Resumo dos tópicos</h2><table class="summary-table"><thead><tr><th>Tópico</th><th>Possível</th><th>Obtido</th><th>Percentual</th></tr></thead><tbody>${topicRows}<tr><th>Total</th><th>${formatPercent(applicableWeight)}</th><th>${formatPercent(checklist.questions.reduce((sum, question) => sum + (question.status === "conforme" ? question.weight : 0), 0))}</th><th>${score} %</th></tr></tbody></table><h2>${reportTitle}</h2>${resultRows}<h2>Observações gerais</h2><div class="notes">${escapeHtml(meta.notes || "Nenhuma observação registrada.")}</div><div class="footer"><span>Desenvolvido por Cherry Checklist · Material de uso interno</span><span>Gerado em ${new Date().toLocaleString("pt-BR")}</span></div><script>window.onload=()=>window.print()<\/script></body></html>`;
  };
  const downloadReport = (content: string, filename: string, delay = 0) => {
    window.setTimeout(() => { const url = URL.createObjectURL(new Blob([content], { type: "text/html;charset=utf-8" })); const link = document.createElement("a"); link.href = url; link.download = filename; link.click(); URL.revokeObjectURL(url); }, delay);
  };
  const saveCompleteReport = async () => {
    if (!isValid) { toast.error("Ajuste os pesos para 100% antes de salvar o relatório."); return; }
    downloadReport(reportHtml("conformidades"), reportFileName("conformidades"));
    downloadReport(reportHtml("nao-conformidades"), reportFileName("nao-conformidades"), 450);
    toast.success("Dois relatórios salvos: conformidades e não conformidades.");
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
    const popup = window.open("", "_blank", "width=980,height=760");
    if (!popup) { toast.error("Permita pop-ups para gerar o PDF."); return; }
    popup.document.write(reportHtml());
    popup.document.close();
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
                    <div className="question-body"><p>{q.text}</p><span className="weight-label">{q.status === "nao-se-aplica" || q.status === "nao-observado" ? "Sem peso contabilizado" : `Peso ${q.weight}%`}</span></div>
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
            <div className="side-card export-card"><div className="side-icon cherry"><Download size={19} /></div><div><h3>Relatórios da avaliação</h3><p>Salve sempre dois arquivos: um de conformidades e outro de não conformidades.</p></div><button className="full-button" onClick={printPdf}><FileDown size={17} /> Visualizar não conformidades</button><button className="outline-button" onClick={saveCompleteReport}><Save size={16} /> Salvar os 2 relatórios</button><button className="outline-button" onClick={shareCompleteReport}><Share2 size={16} /> Enviar arquivo completo</button><button className="outline-button" onClick={startNewReport}><RotateCcw size={16} /> Começar novo relatório</button></div>
            <div className="side-card"><div className="side-icon green"><Share2 size={19} /></div><div><h3>Compartilhar resultado</h3><p>Envie pelo WhatsApp, e-mail ou outro aplicativo.</p></div><button className="outline-button" onClick={share}><Share2 size={17} /> Compartilhar</button><a className="text-link" href={`https://wa.me/?text=${encodeURIComponent(`${checklist.title} — resultado ${score}%`)}`} target="_blank" rel="noreferrer">Abrir no WhatsApp <ChevronRight size={15} /></a><a className="text-link" href={`mailto:?subject=${encodeURIComponent(checklist.title)}&body=${encodeURIComponent(`Resultado: ${score}%`)}`}><Mail size={15} /> Enviar por e-mail <ChevronRight size={15} /></a></div>
            <div className="side-card import-card"><div className="side-icon blue"><Upload size={19} /></div><div><h3>Arquivos de perguntas</h3><p>Digite manualmente ou importe perguntas prontas em DOC/DOCX, TXT, PDF, Excel, CSV ou JSON.</p></div><div className="two-buttons"><button className="outline-button" onClick={loadModelChecklist}><Check size={16} /> Usar checklist Boas Práticas</button><button className="outline-button" onClick={exportQuestionsCsv}><FileUp size={16} /> Excel / CSV</button><button className="outline-button" onClick={exportQuestionsWord}><FileUp size={16} /> Word</button><label className="outline-button"><FileDown size={16} /> Importar<input key={checklist.questions.length} type="file" hidden onChange={(e) => importQuestionsFile(e.target.files?.[0])} /></label></div><div className="file-actions"><button className="text-link" onClick={downloadQuestionTemplate}>Baixar modelo Excel/CSV <ChevronRight size={15} /></button><button className="text-link" onClick={downloadWordTemplate}>Baixar modelo Word: Boas Práticas <ChevronRight size={15} /></button><button className="text-link" onClick={exportJson}>Exportar estrutura completa <ChevronRight size={15} /></button></div></div><div className="side-card app-link-card"><div className="side-icon blue"><Share2 size={19} /></div><div><h3>Enviar para outro celular</h3><p>Compartilhe o link para abrir e instalar o Cherry Checklist em outro aparelho.</p></div><button className="outline-button" onClick={shareAppLink}><Share2 size={17} /> Compartilhar link do app</button><span className="local-note">O link instala o app; os dados permanecem neste celular.</span></div>
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
