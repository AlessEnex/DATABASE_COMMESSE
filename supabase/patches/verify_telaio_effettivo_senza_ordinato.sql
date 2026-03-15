-- PATCH TEMPORANEA (solo verifica, nessuna modifica dati)
-- Scopo:
-- - trovare commesse con data ordine telaio effettiva valorizzata
-- - ma flag telaio_consegnato NON true (false oppure null)

-- Output unico:
-- - elenco commesse incoerenti
-- - colonna "totale_commesse_incoerenti" ripetuta su ogni riga
with incoerenze as (
  select
    c.id,
    c.anno,
    c.numero,
    c.titolo,
    c.cliente,
    c.stato,
    c.tipo_macchina,
    c.variante_macchina,
    c.data_consegna_telaio_effettiva as data_ordine_telaio_effettiva,
    c.telaio_consegnato,
    c.updated_at
  from commesse c
  where c.data_consegna_telaio_effettiva is not null
    and coalesce(c.telaio_consegnato, false) = false
)
select
  i.*,
  count(*) over() as totale_commesse_incoerenti
from incoerenze i
order by i.anno desc, i.numero desc;

-- Se vuoi SOLO il numero totale, usa:
-- select count(*) as totale_commesse_incoerenti
-- from commesse c
-- where c.data_consegna_telaio_effettiva is not null
--   and coalesce(c.telaio_consegnato, false) = false;
