-- Add 'claude-html-fast' to the allowed engines in generated_reports table
ALTER TABLE public.generated_reports
DROP CONSTRAINT IF EXISTS generated_reports_engine_check;

ALTER TABLE public.generated_reports
ADD CONSTRAINT generated_reports_engine_check
CHECK (engine = ANY (ARRAY['gamma'::text, 'claude-pptx'::text, 'pptxgen'::text, 'claude-html'::text, 'claude-html-fast'::text]));
