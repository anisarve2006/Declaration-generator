-- Supabase Table Schema for Declaration Form Generator
-- Run this SQL in Supabase Dashboard -> SQL Editor

CREATE TABLE IF NOT EXISTS public.student_profiles (
    roll_no TEXT PRIMARY KEY,
    student_name TEXT,
    branch TEXT,
    semester TEXT,
    division TEXT,
    last_subject_code TEXT,
    last_subject_name TEXT,
    custom_subjects JSONB DEFAULT '[]'::jsonb,
    signature_data TEXT,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Enable Row Level Security (RLS) & allow anonymous access using ANON_KEY
ALTER TABLE public.student_profiles ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow anon select on student_profiles"
ON public.student_profiles FOR SELECT
USING (true);

CREATE POLICY "Allow anon insert/update on student_profiles"
ON public.student_profiles FOR ALL
USING (true)
WITH CHECK (true);
