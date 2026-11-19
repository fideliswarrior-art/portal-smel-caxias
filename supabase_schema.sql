-- ============================================================================
-- PORTAL ESPORTIVO CAXIAS MAIS ESPORTE (SMEL DUQUE DE CAXIAS)
-- Script de Criação do Banco de Dados + Storage no Supabase (PostgreSQL)
-- Instruções:
-- 1) No painel do Supabase, abra "SQL Editor" > "New query".
-- 2) Cole TODO este arquivo e clique em "Run".
-- ============================================================================

-- 1. TABELAS PRINCIPAIS
CREATE TABLE IF NOT EXISTS public.eventos (
  id TEXT PRIMARY KEY,
  nome TEXT NOT NULL,
  modalidade TEXT DEFAULT '',
  categoria TEXT DEFAULT '',
  local_evento TEXT DEFAULT '',
  data_inicio TEXT DEFAULT '',
  data_fim TEXT DEFAULT '',
  prazo_inscricao TEXT DEFAULT '',
  status TEXT DEFAULT 'Em breve',
  min_atletas TEXT DEFAULT '1',
  max_atletas TEXT DEFAULT '1',
  max_vagas TEXT DEFAULT '',
  descricao TEXT DEFAULT '',
  regulamento_url TEXT DEFAULT '',
  link_challonge_embed TEXT DEFAULT '',
  pasta_fotos_url TEXT DEFAULT '',
  imagem_url TEXT DEFAULT '',
  aceita_individual TEXT DEFAULT '',
  pasta_evento_url TEXT DEFAULT '',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS public.inscricoes (
  id BIGSERIAL PRIMARY KEY,
  data_envio TIMESTAMPTZ DEFAULT NOW(),
  codigo TEXT NOT NULL,
  protocolo TEXT NOT NULL,
  time TEXT DEFAULT '',
  instagram_time TEXT DEFAULT '',
  nome TEXT NOT NULL,
  idade TEXT DEFAULT '',
  telefone TEXT DEFAULT '',
  cpf TEXT NOT NULL,
  email TEXT DEFAULT '',
  rede_social TEXT DEFAULT '',
  status_aprovacao TEXT DEFAULT 'Pendente',
  origem TEXT DEFAULT 'site',
  termo TEXT DEFAULT '',
  evento_id TEXT NOT NULL REFERENCES public.eventos(id) ON DELETE CASCADE,
  evento TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_insc_evento ON public.inscricoes(evento_id);
CREATE INDEX IF NOT EXISTS idx_insc_codigo ON public.inscricoes(codigo);
CREATE INDEX IF NOT EXISTS idx_insc_cpf ON public.inscricoes(cpf);

CREATE TABLE IF NOT EXISTS public.partidas (
  id BIGSERIAL PRIMARY KEY,
  id_jogo TEXT NOT NULL,
  evento_id TEXT NOT NULL REFERENCES public.eventos(id) ON DELETE CASCADE,
  formato TEXT DEFAULT 'mata_mata',
  etapa TEXT DEFAULT '',
  fase TEXT DEFAULT '',
  grupo TEXT DEFAULT '',
  rodada TEXT DEFAULT '',
  ordem TEXT DEFAULT '',
  time_a TEXT DEFAULT '',
  placar_a TEXT DEFAULT '',
  time_b TEXT DEFAULT '',
  placar_b TEXT DEFAULT '',
  vencedor TEXT DEFAULT '',
  status_jogo TEXT DEFAULT 'pendente',
  proximo_jogo_id TEXT DEFAULT '',
  proximo_slot TEXT DEFAULT '',
  origem_a TEXT DEFAULT '',
  origem_b TEXT DEFAULT '',
  num_grupos TEXT DEFAULT '',
  passam_por_grupo TEXT DEFAULT '',
  UNIQUE(evento_id, id_jogo)
);

CREATE INDEX IF NOT EXISTS idx_partidas_evento ON public.partidas(evento_id);

CREATE TABLE IF NOT EXISTS public.colocacoes (
  id BIGSERIAL PRIMARY KEY,
  evento_id TEXT NOT NULL REFERENCES public.eventos(id) ON DELETE CASCADE,
  colocacao TEXT NOT NULL,
  nome_equipe TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS public.feed (
  id BIGSERIAL PRIMARY KEY,
  data TEXT DEFAULT '',
  titulo TEXT NOT NULL,
  legenda TEXT DEFAULT '',
  imagem_url TEXT DEFAULT '',
  categoria TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS public.auditoria (
  id BIGSERIAL PRIMARY KEY,
  data_hora TEXT NOT NULL,
  servidor TEXT NOT NULL,
  acao TEXT NOT NULL,
  detalhes TEXT DEFAULT '',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS public.usuarios (
  id BIGSERIAL PRIMARY KEY,
  login TEXT UNIQUE NOT NULL,
  senha TEXT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Usuário inicial padrão (você pode alterar a senha ou adicionar outros servidores aqui)
INSERT INTO public.usuarios (login, senha)
VALUES ('fidelis', '1234')
ON CONFLICT (login) DO NOTHING;

-- 2. STORAGE BUCKET PARA CAPAS E FOTOS DOS EVENTOS
INSERT INTO storage.buckets (id, name, public)
VALUES ('smel-midias', 'smel-midias', true)
ON CONFLICT (id) DO NOTHING;

-- Políticas de Storage (leitura pública, upload/exclusão liberados para o portal)
DROP POLICY IF EXISTS "Leitura publica midias" ON storage.objects;
CREATE POLICY "Leitura publica midias" ON storage.objects
  FOR SELECT USING (bucket_id = 'smel-midias');

DROP POLICY IF EXISTS "Upload midias portal" ON storage.objects;
CREATE POLICY "Upload midias portal" ON storage.objects
  FOR INSERT WITH CHECK (bucket_id = 'smel-midias');

DROP POLICY IF EXISTS "Delete midias portal" ON storage.objects;
CREATE POLICY "Delete midias portal" ON storage.objects
  FOR DELETE USING (bucket_id = 'smel-midias');

-- 3. HABILITAR RLS E POLÍTICAS DE ACESSO
ALTER TABLE public.eventos ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.inscricoes ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.partidas ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.colocacoes ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.feed ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.auditoria ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.usuarios ENABLE ROW LEVEL SECURITY;

-- Acesso de leitura/escrita via API REST para o portal (os controles de admin são validados na sessão)
CREATE POLICY "Acesso eventos" ON public.eventos FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Acesso inscricoes" ON public.inscricoes FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Acesso partidas" ON public.partidas FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Acesso colocacoes" ON public.colocacoes FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Acesso feed" ON public.feed FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Acesso auditoria" ON public.auditoria FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Acesso usuarios" ON public.usuarios FOR ALL USING (true) WITH CHECK (true);
