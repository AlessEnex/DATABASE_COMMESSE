-- Update v_commesse to include data_arrivo_kit_cavi
-- Drop and recreate to avoid column rename errors

drop view if exists v_commesse;

create view v_commesse as
select
  c.id,
  c.codice,
  c.anno,
  c.numero,
  c.titolo,
  c.cliente,
  c.riferimento,
  c.stato,
  c.stato_op,
  c.priorita,
  c.data_ingresso,
  c.data_consegna_prevista,
  c.data_richiesta,
  c.data_confermata,
  c.ritardo,
  c.conferma_cliente_preliminare3,
  c.conferma_cliente_3d_finale,
  c.data_richiesta_consegna_telaio,
  c.data_conferma_consegna_telaio,
  c.week_consegna_telaio,
  c.ritardo_giorni,
  c.consegna_quadro_richiesta,
  c.quote_pronte,
  c.kit_ordinato,
  c.arrivo_kit,
  c.data_prelievo,
  c.data_consegna_macchina,
  c.trasporto_consegna,
  c.week_ingresso_ordine,
  c.imponibile,
  c.note_generali,
  c.created_at,
  c.updated_at,
  coalesce(
    jsonb_agg(
      jsonb_build_object(
        'reparto_id', r.id,
        'reparto_nome', r.nome,
        'stato', cr.stato,
        'note_reparto', cr.note_reparto,
        'updated_at', cr.updated_at
      )
      order by r.nome
    ) filter (where r.id is not null),
    '[]'::jsonb
  ) as reparti,
  c.data_ordine_telaio,
  c.telaio_consegnato,
  c.data_consegna_telaio_effettiva
from commesse c
left join commesse_reparti cr on cr.commessa_id = c.id
left join reparti r on r.id = cr.reparto_id
group by c.id;
