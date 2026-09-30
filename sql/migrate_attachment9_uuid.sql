-- Run once in MPM_DIC_DEV. Review foreign-key references before running.
-- Keeps each row's data and natural keys; assigns a new UUID to each existing row.
-- Existing numeric IDs are replaced. Run within a maintenance window.

BEGIN;

DO $migration$
DECLARE
    target_table text;
    pk_name text;
    id_kind text;
    incoming_fk text;
BEGIN
    IF current_database() <> 'MPM_DIC_DEV' THEN
        RAISE EXCEPTION 'Wrong database: %', current_database();
    END IF;

    FOREACH target_table IN ARRAY ARRAY[
        'abd_cess_combined_stg',
        'abd_decom_combined_stg',
        'abd_masterplan_combined_stg'
    ] LOOP
        SELECT c.data_type INTO id_kind
        FROM information_schema.columns c
        WHERE c.table_schema = 'amir' AND c.table_name = target_table
          AND c.column_name = 'id';
        IF id_kind IS DISTINCT FROM 'bigint' THEN
            RAISE EXCEPTION 'Expected BIGINT id in amir.%, found %', target_table, id_kind;
        END IF;

        SELECT c.conname INTO pk_name
        FROM pg_constraint c
        WHERE c.conrelid = format('amir.%I', target_table)::regclass
          AND c.contype = 'p'
          AND pg_get_constraintdef(c.oid) = 'PRIMARY KEY (id)';
        IF pk_name IS NULL THEN
            RAISE EXCEPTION 'Expected PRIMARY KEY (id) on amir.%', target_table;
        END IF;

        SELECT c.conname INTO incoming_fk
        FROM pg_constraint c
        WHERE c.confrelid = format('amir.%I', target_table)::regclass
          AND c.contype = 'f'
        LIMIT 1;
        IF incoming_fk IS NOT NULL THEN
            RAISE EXCEPTION 'amir.% is referenced by foreign key %; migrate the referencing table first',
                target_table, incoming_fk;
        END IF;

        EXECUTE format('ALTER TABLE amir.%I ADD COLUMN id_uuid uuid DEFAULT gen_random_uuid() NOT NULL', target_table);
        EXECUTE format('ALTER TABLE amir.%I DROP CONSTRAINT %I', target_table, pk_name);
        EXECUTE format('ALTER TABLE amir.%I DROP COLUMN id', target_table);
        EXECUTE format('ALTER TABLE amir.%I RENAME COLUMN id_uuid TO id', target_table);
        EXECUTE format('ALTER TABLE amir.%I ADD CONSTRAINT %I PRIMARY KEY (id)', target_table, pk_name);
    END LOOP;
END
$migration$;

COMMIT;

SELECT table_name, data_type, column_default, is_nullable
FROM information_schema.columns
WHERE table_schema = 'amir'
  AND table_name IN ('abd_cess_combined_stg', 'abd_decom_combined_stg',
                     'abd_masterplan_combined_stg')
  AND column_name = 'id'
ORDER BY table_name;
