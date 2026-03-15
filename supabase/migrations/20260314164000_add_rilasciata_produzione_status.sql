-- Add explicit commessa status for production release.

do $$
begin
  if exists (select 1 from pg_type where typname = 'commessa_stato') then
    if not exists (
      select 1
      from pg_enum
      where enumtypid = 'commessa_stato'::regtype
        and enumlabel = 'rilasciata_produzione'
    ) then
      alter type commessa_stato add value 'rilasciata_produzione' after 'in_corso';
    end if;
  end if;
end $$;
