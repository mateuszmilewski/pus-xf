Attribute VB_Name = "Module4"
Public perimetre_analyse, lig_cofor_carre, num_shape, test_ref_style, domaine_s, lig_ref_ac, nom_fnr_a, index_fnr, index_noa, index_noa_fnr, reference_ac
Public Debut, InitSB, tab_sechel, feuille_courante, i, j, k, m, x, nbre_ligne_filtree, valeur1, n1, n2, n3, nb_lig_noa, nb_lig_noa_fnr, nb_lig_fnr, info_dap
Public prog_livraison, contrat, ref_affect, noa, der_lig_ref_s, reference, nom_contact_log, tel_contact_log, mail_contact_log, nom_fnr_b, nom_contact_b, nom_cofor_exp_b
Public echeance, tmc, qte_conf, num_tmc, noa_a, cofor_a, reference_a, lig_ref_p, ref_tmc, cofor_fnr_s, nom_fnr_s, confirm, nb_confirm_non, nb_confirm_oui, nom_appro, ligne_noa_fnr, ligne_deroul
'Public col_base_commSechel, col_base_imputationCodeI, col_base_commRPO, col_base_VOR, col_base_acteur


'parametrage onglet BASE - colonnes ajoutées en violet
Public Property Get col_base_commSechel() As Variant
col_base_commSechel = 49
End Property
Public Property Get col_base_imputationCodeI() As Variant
col_base_imputationCodeI = 50
End Property
Public Property Get col_base_commRPO() As Variant
col_base_commRPO = 51
End Property
Public Property Get col_base_VOR() As Variant
col_base_VOR = 52
End Property
Public Property Get col_base_acteur() As Variant
col_base_acteur = 53
End Property
Public Property Get col_base_ech() As Variant
col_base_ech = 54
End Property
Public Property Get col_base_mag() As Variant
col_base_mag = 55
End Property
Public Property Get col_base_fnr() As Variant
col_base_fnr = 56
End Property


'nb colonne PUS
Public Property Get lastColPUS() As Variant
lastColPUS = 56
End Property

'parametrage onglet BASE
Public Property Get prem_lig_ref_b() As Variant
prem_lig_ref_b = 3
End Property
Public Property Get col_ref_b() As Variant
col_ref_b = 1
End Property
Public Property Get col_desi_b() As Variant
col_desi_b = 2
End Property
Public Property Get col_qte_theo_b() As Variant
col_qte_theo_b = 3
End Property
Public Property Get col_qte_conf_b() As Variant
col_qte_conf_b = 4
End Property

Public Property Get col_cofor_b() As Variant
col_cofor_b = 6
End Property
Public Property Get col_nom_fnr_b() As Variant
col_nom_fnr_b = 7
End Property
Public Property Get col_del_prog_b() As Variant
col_del_prog_b = 8
End Property
Public Property Get col_po_numb_b() As Variant
col_po_numb_b = 9
End Property
Public Property Get col_nom_appro_b() As Variant
col_nom_appro_b = 10
End Property
Public Property Get col_tel_appro_b() As Variant
col_tel_appro_b = 11
End Property
Public Property Get col_mail_appro_b() As Variant
col_mail_appro_b = 12
End Property
Public Property Get col_nom_log_b() As Variant
col_nom_log_b = 13
End Property
Public Property Get col_tel_log_b() As Variant
col_tel_log_b = 14
End Property
Public Property Get col_mail_log_b() As Variant
col_mail_log_b = 15
End Property
Public Property Get col_tmc_b() As Variant
col_tmc_b = 16
End Property
Public Property Get col_cofor_vend_b() As Variant
col_cofor_vend_b = 17
End Property
Public Property Get col_cofor_exp_b() As Variant
col_cofor_exp_b = 18
End Property
Public Property Get col_pickup_address_b() As Variant
col_pickup_address_b = 20
End Property
Public Property Get col_perimetre_b() As Variant
col_perimetre_b = 23
End Property
Public Property Get col_pack_id_b() As Variant
col_pack_id_b = 24
End Property
Public Property Get col_pack_amount_b() As Variant
col_pack_amount_b = 46
End Property
Public Property Get col_qty_in_pack_b() As Variant
col_qty_in_pack_b = 25
End Property
Public Property Get col_pack_lo_b() As Variant
col_pack_lo_b = 27
End Property
Public Property Get col_pack_la_b() As Variant
col_pack_la_b = 28
End Property
Public Property Get col_pack_ha_b() As Variant
col_pack_ha_b = 29
End Property
Public Property Get col_pack_weig_b() As Variant
col_pack_weig_b = 26
End Property
Public Property Get col_hazard_b() As Variant
col_hazard_b = 31
End Property
Public Property Get col_stack_b() As Variant
col_stack_b = 32
End Property
Public Property Get col_alt_pack_id_b() As Variant
col_alt_pack_id_b = 33
End Property
Public Property Get col_alt_pack_weig_b() As Variant
col_alt_pack_weig_b = 34
End Property
Public Property Get col_alt_pack_lo_b() As Variant
col_alt_pack_lo_b = 35
End Property
Public Property Get col_alt_pack_la_b() As Variant
col_alt_pack_la_b = 36
End Property
Public Property Get col_alt_pack_ha_b() As Variant
col_alt_pack_ha_b = 37
End Property
Public Property Get col_pickup_date_b() As Variant
col_pickup_date_b = 44
End Property
Public Property Get col_psa_contact_1_b() As Variant
col_psa_contact_1_b = 10
End Property
Public Property Get col_psa_contact_2_b() As Variant
col_psa_contact_2_b = 11
End Property
Public Property Get col_psa_contact_3_b() As Variant
col_psa_contact_3_b = 12
End Property
Public Property Get col_psa_supply_1_b() As Variant
col_psa_supply_1_b = 13
End Property
Public Property Get col_psa_supply_2_b() As Variant
col_psa_supply_2_b = 14
End Property
Public Property Get col_psa_supply_3_b() As Variant
col_psa_supply_3_b = 15
End Property

Public Property Get col_gefco_1_b() As Variant
col_gefco_1_b = 39
End Property
Public Property Get col_gefco_2_b() As Variant
col_gefco_2_b = 21
End Property
Public Property Get col_gefco_3_b() As Variant
col_gefco_3_b = 47
End Property
Public Property Get col_gefco_4_b() As Variant
col_gefco_4_b = 40
End Property
Public Property Get col_gefco_5_b() As Variant
col_gefco_5_b = 43
End Property
Public Property Get col_gefco_6_b() As Variant
col_gefco_6_b = 42
End Property

Public Property Get prem_lig_ref_cpl() As Variant
prem_lig_ref_cpl = 3
End Property
Public Property Get col_cpl_prem() As Variant
col_cpl_prem = 17
End Property
Public Property Get col_cpl_der() As Variant
col_cpl_der = 48
End Property
'parametrage accueil
Public Property Get lig_cofor_rond() As Variant
lig_cofor_rond = 3
End Property
Public Property Get prem_lig_ref_ac() As Variant
prem_lig_ref_ac = 7
End Property
Public Property Get col_noa_ac() As Variant
col_noa_ac = 2
End Property
Public Property Get col_ref_ac() As Variant
col_ref_ac = 3
End Property
Public Property Get col_desi_ac() As Variant
col_desi_ac = 4
End Property
Public Property Get col_cofor_ac() As Variant
col_cofor_ac = 5
End Property
Public Property Get col_fnr_ac() As Variant
col_fnr_ac = 6
End Property
Public Property Get col_prog_liv_ac() As Variant
col_prog_liv_ac = 7
End Property
'parametrage sechel
Public Property Get prem_lig_ref_s() As Variant
prem_lig_ref_s = 2
End Property
Public Property Get col_noa_s() As Variant
col_noa_s = 1
End Property
Public Property Get col_ref_s() As Variant
col_ref_s = 2
End Property
Public Property Get col_desi_s() As Variant
col_desi_s = 3
End Property
Public Property Get col_cofor_s() As Variant
col_cofor_s = 4
End Property
Public Property Get col_fnr_s() As Variant
col_fnr_s = 5
End Property
Public Property Get col_ech_s() As Variant
col_ech_s = 12
End Property
Public Property Get col_prog_liv_s() As Variant
col_prog_liv_s = 8
End Property
'parametrage pickup sheet
Public Property Get prem_lig_ref_p() As Variant
prem_lig_ref_p = 27
End Property
Public Property Get lig_nom_fnr_p() As Variant
lig_nom_fnr_p = 8
End Property
Public Property Get lig_cofor_vend_p() As Variant
lig_cofor_vend_p = 9
End Property
Public Property Get lig_cofor_exp_p() As Variant
lig_cofor_exp_p = 10
End Property
Public Property Get lig_pickup_address_p() As Variant
lig_pickup_address_p = 11
End Property
Public Property Get lig_psa_contact_1_p() As Variant
lig_psa_contact_1_p = 12
End Property
Public Property Get lig_psa_contact_2_p() As Variant
lig_psa_contact_2_p = 13
End Property
Public Property Get lig_psa_contact_3_p() As Variant
lig_psa_contact_3_p = 14
End Property
Public Property Get lig_psa_supply_1_p() As Variant
lig_psa_supply_1_p = 15
End Property
Public Property Get lig_psa_supply_2_p() As Variant
lig_psa_supply_2_p = 16
End Property
Public Property Get lig_psa_supply_3_p() As Variant
lig_psa_supply_3_p = 17
End Property
Public Property Get lig_psa_dispatch_1_p() As Variant
lig_psa_dispatch_1_p = 18
End Property
Public Property Get lig_psa_dispatch_2_p() As Variant
lig_psa_dispatch_2_p = 19
End Property
Public Property Get lig_psa_dispatch_3_p() As Variant
lig_psa_dispatch_3_p = 20
End Property
Public Property Get lig_gefco_contact_p() As Variant
lig_gefco_contact_p = 25
End Property
Public Property Get lig_gefco_cofor_p() As Variant
lig_gefco_cofor_p = 26
End Property
Public Property Get col_index_p() As Variant
col_index_p = 1
End Property
Public Property Get col_ref_p() As Variant
col_ref_p = 2
End Property
Public Property Get col_desi_p() As Variant
col_desi_p = 3
End Property
Public Property Get col_qte_p() As Variant
col_qte_p = 4
End Property
Public Property Get col_pick_week_p() As Variant
col_pick_week_p = 5
End Property
Public Property Get col_del_prog_p() As Variant
col_del_prog_p = 7
End Property
Public Property Get col_po_numb_p() As Variant
col_po_numb_p = 8
End Property
Public Property Get col_pack_id_p() As Variant
col_pack_id_p = 9
End Property
Public Property Get col_pack_amount_p() As Variant
col_pack_amount_p = 10
End Property
Public Property Get col_qty_in_pack_p() As Variant
col_qty_in_pack_p = 11
End Property
Public Property Get col_pack_dim_p() As Variant
col_pack_dim_p = 12
End Property
Public Property Get col_pack_weig_p() As Variant
col_pack_weig_p = 13
End Property
Public Property Get col_hazard_p() As Variant
col_hazard_p = 14
End Property
Public Property Get col_stack_p() As Variant
col_stack_p = 15
End Property
Public Property Get col_pickup_date_p() As Variant
col_pickup_date_p = 16
End Property
Public Property Get col_alt_pack_dim_p() As Variant
col_alt_pack_dim_p = 17
End Property
Public Property Get col_alt_pack_weig_p() As Variant
col_alt_pack_weig_p = 18
End Property
Public Property Get col_alt_pack_id_p() As Variant
col_alt_pack_id_p = 19
End Property

