"""
referentiel_indicateurs.py - Matrice maîtresse des indicateurs CGPP (riposte MVE-17).

Une seule liste, des identifiants stables en texte (PILIER-NN), et pour chaque indicateur :
la correspondance avec la matrice de Dr Barclay (plan DOS), le cadre global (25 indicateurs),
les cibles DOS et Tranche 5, la source primaire et le champ de la fiche SBC/ODK.

Piliers : COORD, SURV, CREC, PCI, LOG, MEL (structure du plan de travail DOS)
          + PREP, DEPL (piliers propres à la Tranche 5, à activer si elle est retenue).
Statut  : existant / à corriger / nouveau.
Priorité: 1 = à mettre en place immédiatement ; 2 = dans le mois ; 3 = trimestre.
"""

COLS = ["id_indicateur", "pilier", "libelle_fr", "libelle_en", "type", "niveau", "numerateur",
        "denominateur", "source_donnees", "champ_source", "frequence", "desagregation", "cible_dos",
        "cible_tranche5", "seuil_alerte", "id_barclay", "id_cadre_global", "statut", "priorite",
        "commentaire"]

SBC = "Fiche SBC journalière (ODK)"
LL = "Liste linéaire des cas (DHIS2 Tracker) et liste des contacts"
SIT = "SitRep national (tableaux par ZS)"
H, M, Q, J = "Hebdomadaire", "Mensuelle", "Trimestrielle", "Quotidienne"
ZSIP = "IP, ZS"

I = []  # liste des indicateurs


def ind(id_, pilier, fr, en, type_, niveau, num, den, source, champ, freq, desag, c_dos, c_t5,
        seuil, id_b, id_g, statut, prio, comm=""):
    I.append(dict(zip(COLS, [id_, pilier, fr, en, type_, niveau, num, den, source, champ, freq, desag,
                             c_dos, c_t5, seuil, id_b, id_g, statut, prio, comm])))


# ----------------------------------------------------------------- COORDINATION
ind("COORD-01", "COORD", "Matrices 4W élaborées et mises à jour", "# of 4W matrices developed and updated",
    "Nombre", "Tier 2", "Matrices 4W (1 nationale, 3 provinciales) mises à jour dans le mois", "N/A",
    "Matrices 4W", "", M, "Niveau (national, provincial)", "4", "", "", "1.1", "", "existant", 2)
ind("COORD-02", "COORD", "Matrice des partenaires CGPP tenue à jour", "# of CGPP partner matrices maintained",
    "Nombre", "Tier 2", "Matrices partenaires mises à jour", "N/A", "Matrice partenaires", "", M, "",
    "1", "", "", "1.2", "", "existant", 3)
ind("COORD-03", "COORD", "Réunions de coordination auxquelles CGPP a participé ou qu'il a appuyées",
    "# of coordination meetings participated in or supported", "Nombre", "Tier 2",
    "Réunions avec compte rendu et liste de présence complète", "N/A", "Feuille « reunions » de la base",
    "reunions", M, "Niveau, province", "90", "", "", "1.3", "21", "à corriger", 1,
    "Ne compter que les réunions documentées (compte rendu + liste de présence complète).")
ind("COORD-04", "COORD", "Réunions bilatérales ou conjointes avec les partenaires financés par le DOS",
    "# of bilateral/joint coordination meetings with DOS-funded partners", "Nombre", "Tier 2",
    "Réunions bilatérales/conjointes tenues", "N/A", "Notes de réunion", "reunions", M, "Partenaire",
    "12", "", "", "1.4", "", "existant", 3)
ind("COORD-05", "COORD", "Réunions hebdomadaires de coordination CGPP", "# of weekly CGPP coordination meetings conducted",
    "Nombre", "Tier 2", "Réunions internes hebdomadaires tenues", "N/A", "Comptes rendus", "reunions", H, "",
    "16", "", "", "1.5", "", "existant", 3)
ind("COORD-06", "COORD", "Réunions de revue provinciales", "# of provincial review meetings conducted", "Nombre",
    "Tier 2", "Revues provinciales tenues", "N/A", "Rapports de revue", "reunions", M, "Province", "12", "",
    "", "1.6", "", "existant", 3)
ind("COORD-07", "COORD", "Outil harmonisé de suivi EIR et de rapportage mis en place",
    "# of RRT/data reporting trackers established", "Nombre", "Tier 2", "Outil en service", "N/A",
    "Outil de suivi", "", M, "", "1", "", "", "1.7", "", "existant", 1,
    "La base Excel normalisée et la chaîne Python décrites dans ce rapport peuvent tenir ce rôle.")
ind("COORD-08", "COORD", "Mises à jour quotidiennes de terrain transmises", "# of daily field updates submitted",
    "Nombre", "Tier 2", "Mises à jour quotidiennes transmises", "N/A", "Journal des envois", "", J, "IP",
    "110", "", "", "1.8", "", "existant", 2)
ind("COORD-09", "COORD", "Mises à jour hebdomadaires de coordination transmises",
    "# of weekly coordination updates submitted", "Nombre", "Tier 2", "Mises à jour hebdomadaires transmises",
    "N/A", "Journal des envois", "", H, "", "16", "", "", "1.9", "", "à corriger", 2,
    "Le plan DOS indique 16 (extrants du pilier 1) et 18 (activité 1.5) : arrêter une seule cible.")
ind("COORD-10", "COORD", "Séances de négociation ou de médiation de crise avec le pilier CREC",
    "# of crisis negotiation/mediation sessions conducted with CREC Pillar", "Nombre", "Tier 2",
    "Séances tenues", "N/A", "Rapports de négociation", "", "Selon besoin", "", "4", "", "", "1.10", "",
    "existant", 3)
ind("COORD-11", "COORD", "Réunions de coordination transfrontalière", "# of cross-border coordination meetings held",
    "Nombre", "Tier 2", "Réunions transfrontalières tenues", "N/A", "Comptes rendus", "reunions", M, "Frontière",
    "6", "100 échanges transfrontaliers ou interprovinciaux", "", "1.11", "", "existant", 3)
ind("COORD-12", "COORD", "Points d'action des réunions réalisés dans les délais",
    "% of action points generated and followed up", "Pourcentage", "Tier 1",
    "Points d'action réalisés à l'échéance", "Points d'action arrivés à échéance",
    "Feuille « points_action » de la base", "points_action", M, "Niveau de réunion", "100 %", "",
    "< 70 % : rouge", "1.12", "", "à corriger", 1,
    "Intitulé « # » alors que c'est un pourcentage ; cible codée 1 ; les comptes rendus actuels ne listent pas de points d'action.")

# ----------------------------------------------------------------- SURVEILLANCE
ind("SURV-01", "SURV", "Contacts listés et suivis", "# of contacts line-listed and monitored", "Nombre", "Tier 1",
    "Contacts listés (cas confirmés, probables, suspects) entrés dans le système de suivi", "N/A",
    "Liste des contacts ; " + SBC, "contacts_listes_h + contacts_listes_f", H, "Sexe, âge, ZS, niveau de risque",
    "2 000", "30 000", "", "2.1", "8", "existant", 1)
ind("SURV-02", "SURV", "Contacts suivis chaque jour pendant les 21 jours",
    "% of listed contacts followed daily during 21-day follow-up", "Pourcentage", "Tier 1",
    "Contacts-jours vus", "Contacts-jours attendus (contacts sous suivi x jours)", SBC,
    "contacts_vus / contacts_sous_suivi", H, "ZS, sexe", "≥ 90 %", "≥ 90 %", "< 90 % : rouge", "2.2", "9",
    "à corriger", 1,
    "La définition actuelle mélange le suivi quotidien et l'issue à J21 : l'issue devient SURV-21.")
ind("SURV-03", "SURV", "Investigations de cas réalisées", "# of case investigations conducted", "Nombre",
    "Tier 2", "Investigations (cas confirmés, probables, suspects) appuyées", "N/A", "Rapports d'investigation",
    "", H, "ZS, classification", "300", "2 000", "", "2.3", "", "existant", 2)
ind("SURV-04", "SURV", "Cartographies d'exposition réalisées", "# of exposure mappings conducted", "Nombre",
    "Tier 2", "Cartographies d'exposition autour d'un cas", "N/A", "Rapports de cartographie", "", H, "ZS",
    "300", "Incluses dans les 2 000 investigations", "", "2.4", "", "existant", 3)
ind("SURV-05", "SURV", "Alertes de suspicion de MVE détectées par les systèmes appuyés par CGPP",
    "# of suspected EVD alerts detected through CGPP-supported systems", "Nombre", "Tier 1",
    "Alertes (malades et décès) détectées par les canaux appuyés par CGPP", "N/A", SBC,
    "alertes_malades + alertes_deces", H, "Type (malade/décès), source, ZS", "1 500", "35 000", "", "2.5", "1",
    "existant", 1)
ind("SURV-06", "SURV", "Alertes CGPP vérifiées en 24 h", "% of suspected EVD alerts reported by CGPP verified within 24 hours",
    "Pourcentage", "Tier 1", "Alertes vérifiées par l'ESS/EIR dans les 24 h", "Alertes reçues", SBC,
    "alertes_verifiees_24h (nouveau champ) / alertes", H, "ZS", "≥ 80 %", "≥ 85 %", "< 80 % : rouge", "2.6", "3",
    "à corriger", 1, "Aucun champ ne mesure la vérification dans la fiche SBC actuelle.")
ind("SURV-07", "SURV", "Alertes notifiées aux autorités en 24 h", "% of suspected EVD alerts notified to authorities within 24 hours",
    "Pourcentage", "Tier 1", "Alertes signalées au centre de santé dans les 24 h", "Alertes détectées", SBC,
    "alertes_*_notifiees_24h / alertes", H, "ZS", "100 %", "", "< 90 % : rouge", "2.7", "4", "à corriger", 1,
    "Cible codée « 1 » dans la matrice : à écrire 100 %.")
ind("SURV-08", "SURV", "Alertes investiguées dans les délais",
    "% of suspected EVD alerts investigated within 7 days", "Pourcentage", "Tier 1",
    "Alertes investiguées dans le délai retenu", "Alertes détectées", SBC, "alertes_investiguees_24h", H, "ZS",
    "100 %", "", "", "2.8", "5", "à corriger", 1,
    "Le cadre dit 7 jours, la fiche SBC mesure 24 h : retenir 24 h (cohérent avec 7-1-7 et le SitRep).")
ind("SURV-09", "SURV", "Acteurs de la surveillance à base communautaire formés et activés",
    "# of CBS actors/community volunteers trained and activated", "Nombre", "Tier 2",
    "RECO, relais, informateurs, leaders formés et actifs", "N/A", "Rapports de formation ; listes d'activation",
    "", M, "Sexe, type d'acteur, ZS", "2 024", "12 000", "", "2.9", "6", "existant", 2)
ind("SURV-10", "SURV", "Exercices PopCAB ou de cartographie des risques de mobilité",
    "# of risk mapping exercises conducted", "Nombre", "Tier 2", "Exercices PopCAB réalisés", "N/A",
    "Rapports PopCAB", "", M, "ZS", "12", "60", "", "2.10", "7", "à corriger", 3,
    "Fréquence « annuelle » incompatible avec une cible de 12 : passer en mensuel.")
ind("SURV-11", "SURV", "Alertes communautaires reçues", "# of community alerts received", "Nombre", "Tier 2",
    "Alertes communautaires et de FOSA reçues", "N/A", "Registre des alertes", "", H, "Source", "1 500", "",
    "", "2.11", "", "à corriger", 2, "Même cible et définition voisine de SURV-05 : fusionner ou distinguer clairement.")
ind("SURV-12", "SURV", "Déploiements d'équipes d'intervention rapide (EIR) appuyés",
    "# of RRT deployments supported", "Nombre", "Tier 2", "Déploiements EIR appuyés", "N/A", "Journal des EIR", "",
    H, "ZS", "300", "500 membres d'EIR recyclés", "", "2.12", "23", "existant", 2,
    "L'indicateur 23 du cadre global mélange équipes, déploiements et alertes : il est scindé en SURV-12 et LOG-02.")
ind("SURV-13", "SURV", "PoE informels, PoC et points de mobilité appuyés pour le dépistage",
    "# of informal PoEs/PoCs supported for screening", "Nombre", "Tier 2", "Sites appuyés et fonctionnels",
    "N/A", "Évaluations de site", "", M, "Type de site", "60", "150", "", "2.13", "17", "existant", 2)
ind("SURV-14", "SURV", "Personnes dépistées aux PoE/PoC ou points de mobilité",
    "# of people screened at PoEs/PoCs or mobility points", "Nombre", "Tier 1", "Voyageurs dépistés", "N/A",
    "Registres de dépistage", "", H, "Site, sexe", "269 518", "", "", "2.14", "16", "existant", 2)
ind("SURV-15", "SURV", "Réseaux d'alerte frontaliers et communautaires établis",
    "# of border/community alert networks established", "Nombre", "Tier 2", "Réseaux activés", "N/A",
    "Rapports d'activation", "", M, "Frontière", "20", "", "", "2.15", "", "existant", 3)
ind("SURV-16", "SURV", "Alertes et références transfrontalières documentées et suivies",
    "# of cross-border alerts/referrals documented and followed up", "Nombre", "Tier 2", "Alertes suivies",
    "N/A", "Journal de notification", "", H, "Pays", "70", "500", "", "2.16", "", "existant", 2)
ind("SURV-17", "SURV", "Cas confirmés dans les zones d'intervention CGPP",
    "# of confirmed Ebola cases identified within CGPP focal areas", "Nombre", "Contexte",
    "Cas confirmés (toutes sources) dans les ZS CGPP", "N/A", SIT, "epi_zs.cas_cumul, nouveaux_cas_24h", J,
    "ZS, IP", "Suivi", "", "", "2.17", "2", "à corriger", 1,
    "À calculer automatiquement depuis le SitRep national et non à déclarer par les IP (le classeur affiche 1, le SitRep 2 586 au 10/10).")
ind("SURV-18", "SURV", "Décès inexpliqués ou communautaires investigués",
    "# of unexplained/community deaths investigated", "Nombre", "Tier 2", "Décès investigués", "N/A",
    "Journal des décès", "deces_communautaires_notifies", H, "ZS", "50", "", "", "2.18", "", "existant", 2)
ind("SURV-19", "SURV", "Contacts symptomatiques référés vers une FOSA ou un CTE",
    "# of symptomatic contacts referred to health facilities/ETC", "Nombre", "Tier 2", "Contacts référés",
    "N/A", SBC, "contacts_symptomatiques_referes_24h", H, "ZS", "100", "", "", "2.19", "", "existant", 2)
ind("SURV-20", "SURV", "Nouveaux cas confirmés qui étaient des contacts connus",
    "% of new confirmed cases who were known contacts", "Pourcentage", "Tier 1",
    "Nouveaux cas confirmés figurant sur une liste de contacts avant leur confirmation",
    "Nouveaux cas confirmés dans les ZS CGPP", LL, "", H, "ZS", "", "",
    "Cible proposée ≥ 80 % (à valider)", "", "", "nouveau", 1,
    "Mesure directe de l'efficacité du suivi des contacts : c'est l'indicateur qui manque le plus.")
ind("SURV-21", "SURV", "Contacts ayant terminé les 21 jours de suivi",
    "% of contacts completing 21-day follow-up", "Pourcentage", "Tier 1",
    "Contacts sortis du suivi à J21 sans interruption", "Contacts dont la fenêtre de 21 jours s'est terminée",
    SBC, "contacts_fin_suivi_21j (nouveau)", H, "ZS", "", "", "< 90 % : rouge", "", "", "nouveau", 1,
    "Publier aussi le pourcentage de perdus de vue (champ contacts_perdus_de_vue).")
ind("SURV-22", "SURV", "Nouveaux cas confirmés détectés après le décès (décès communautaires)",
    "% of new confirmed cases detected post-mortem", "Pourcentage", "Tier 1",
    "Nouveaux cas confirmés issus de décès communautaires (swab +)", "Nouveaux cas confirmés", SIT,
    "epi_zs.deces_communautaires_24h / nouveaux_cas_24h", H, "ZS", "", "", "Tendance à la baisse attendue",
    "", "", "nouveau", 1, "Calculable dès maintenant : 44 % le 07/10 et 20 % le 10/10.")
ind("SURV-23", "SURV", "Délai entre le début des symptômes et l'isolement",
    "Median delay from symptom onset to isolation (days)", "Délai", "Tier 1",
    "Médiane des délais (jours) des cas confirmés de la semaine ; part isolée en 3 jours ou moins", "N/A", LL, "",
    H, "ZS", "", "", "Médiane > 3 j : rouge (seuil à valider)", "", "", "nouveau", 2)
ind("SURV-24", "SURV", "Respect des délais 7-1-7 pour les événements détectés",
    "% of events meeting 7-1-7 timeliness targets", "Pourcentage", "Tier 1",
    "Événements détectés ≤ 7 j, notifiés ≤ 1 j, réponse initiale ≤ 7 j", "Événements validés",
    "Registre des alertes et événements", "", M, "Étape (détection, notification, réponse)", "", "", "",
    "", "", "nouveau", 2, "Le plan DOS prévoit l'intégration du suivi 7-1-7 (activité 1.7.6 de la Tranche 5).")
ind("SURV-25", "SURV", "Contacts listés par cas confirmé", "Mean number of contacts listed per confirmed case",
    "Ratio", "Tier 1", "Contacts listés autour des cas de la semaine", "Cas confirmés de la semaine", LL, "", H,
    "ZS", "", "", "", "", "", "nouveau", 2, "Un ratio bas signale un listage incomplet.")
ind("SURV-26", "SURV", "Contacts symptomatiques référés en 24 h ou moins",
    "% of symptomatic contacts referred within 24 hours", "Pourcentage", "Tier 1",
    "Contacts symptomatiques référés dans les 24 h", "Contacts ayant présenté au moins un signe", SBC,
    "contacts_symptomatiques_referes_24h / contacts_symptomatiques", H, "ZS", "", "", "< 90 % : rouge", "", "",
    "nouveau", 1)
ind("SURV-27", "SURV", "Alertes CGPP validées comme cas suspects (valeur prédictive)",
    "% of CGPP alerts validated as suspected cases", "Pourcentage", "Tier 2", "Alertes CGPP validées",
    "Alertes CGPP vérifiées", "Registre des alertes", "", H, "ZS, type", "", "", "", "", "", "nouveau", 2,
    "Mesure la qualité des alertes communautaires (au niveau national : 15,8 % le 07/10).")
ind("SURV-28", "SURV", "Alertes pour 10 000 habitants par semaine", "Alert rate per 10,000 population per week",
    "Ratio", "Tier 2", "Alertes de la semaine x 10 000", "Population de la ZS", SBC + " ; ref_zones.population",
    "", H, "ZS", "", "", "", "", "", "nouveau", 2, "Sensibilité de la surveillance communautaire ; nécessite la population par ZS.")
ind("SURV-29", "SURV", "Complétude des rapports SBC journaliers", "% of expected CBS daily reports received",
    "Pourcentage", "Tier 2", "RECO/ACSA ayant travaillé et rapporté", "RECO/ACSA attendus", SBC,
    "reco_ayant_travaille / reco_attendus", J, "AS, ZS, IP", "", "", "< 80 % : rouge", "", "", "nouveau", 1,
    "Les champs existent déjà dans l'en-tête de la fiche SBC.")
ind("SURV-30", "SURV", "Décès communautaires notifiés et prélevés (swab)",
    "% of notified community deaths swabbed", "Pourcentage", "Tier 1", "Décès communautaires prélevés",
    "Décès communautaires notifiés", SBC, "deces_communautaires_swabes / deces_communautaires_notifies", H, "ZS",
    "", "", "< 90 % : rouge", "", "", "nouveau", 1)
ind("SURV-31", "SURV", "Alertes de santé animale (One Health) transmises aux services vétérinaires",
    "# of animal health alerts relayed (One Health)", "Nombre", "Tier 2", "Alertes animales transmises", "N/A",
    SBC, "alertes_sante_animale (nouveau)", H, "ZS", "", "", "", "", "", "nouveau", 3)

# ----------------------------------------------------------------- CREC
ind("CREC-01", "CREC", "Ménages visités et séances d'engagement communautaire",
    "# of households reached with RCCE messages", "Nombre", "Tier 2",
    "Ménages visités (porte à porte) ; séances d'engagement (à rapporter séparément)", "N/A", SBC,
    "nb_menages_visites ; nb_seances_realisees", M, "Type (ménage, séance), ZS", "2 515", "40 000", "",
    "3.1", "10", "à corriger", 1,
    "L'intitulé parle de ménages mais le numérateur compte des séances : scinder en deux chiffres.")
ind("CREC-02", "CREC", "Séances de groupe ciblées", "# of targeted RCCE group sessions conducted", "Nombre",
    "Tier 2", "Dialogues, forums, causeries, réunions de groupe", "N/A", SBC,
    "nb_causeries + nb_debats_dialogues (hors VAD)", M, "Technique, ZS", "300", "", "", "3.2", "11", "existant", 2)
ind("CREC-03", "CREC", "Personnes atteintes par les messages CREC", "# of people reached with EVD RCCE messaging",
    "Nombre", "Tier 1", "Personnes sensibilisées", "N/A", SBC, "sensibilises_h + sensibilises_f", M,
    "Sexe, canal (VAD, groupe, médias), ZS", "550 810", "5 000 000", "", "3.3", "12", "existant", 1,
    "Ne pas additionner l'audience radio estimée et les personnes atteintes en face à face.")
ind("CREC-04", "CREC", "Séances ciblées en milieux à haut risque", "# of targeted RCCE sessions conducted in high-risk settings",
    "Nombre", "Tier 2", "Séances dans les sites à haut risque de la liste fixée", "N/A", SBC,
    "nb_seances_realisees où type_site ∈ {école, marché, site minier, PoE, déplacés, parkings, guérison trad.}",
    M, "Type de site", "300", "", "", "3.4", "13", "à corriger", 2, "Fixer la liste des types de site « à haut risque ».")
ind("CREC-05", "CREC", "Rumeurs et préoccupations communautaires enregistrées",
    "# of rumors or community concerns logged", "Nombre", "Tier 2", "Rumeurs/préoccupations enregistrées", "N/A",
    SBC, "rumeurs_enregistrees (nouveau)", H, "Catégorie, ZS", "500", "10 000", "", "3.5", "14", "existant", 1)
ind("CREC-06", "CREC", "Rumeurs et préoccupations traitées en 48 h",
    "% of rumors/community concerns addressed within 48 hours", "Pourcentage", "Tier 1",
    "Rumeurs traitées dans les 48 h", "Rumeurs enregistrées", SBC, "rumeurs_traitees_48h / rumeurs_enregistrees",
    H, "Catégorie", "≥ 80 %", "≥ 80 %", "< 80 % : rouge", "3.6", "15", "à corriger", 1,
    "La fiche SBC mesure des réticences levées le jour même : ajouter les champs rumeurs.")
ind("CREC-07", "CREC", "Supports IEC produits ou distribués", "# of IEC materials produced/distributed", "Nombre",
    "Tier 2", "Supports produits/distribués", "N/A", "Registres de production", "", M, "Type, langue", "150 000",
    "1 500 000", "", "3.7", "", "existant", 3)
ind("CREC-08", "CREC", "Spots ou jingles radio diffusés", "# of radio spots/jingles aired", "Nombre", "Tier 2",
    "Spots diffusés", "N/A", "Journaux de diffusion", "nb_spots_radio", M, "Station", "400", "3 000", "", "3.8", "",
    "existant", 3)
ind("CREC-09", "CREC", "Émissions radio interactives", "# of interactive radio programs conducted", "Nombre",
    "Tier 2", "Émissions réalisées", "N/A", "Journaux de diffusion", "nb_emissions_interactives", M, "Station", "24",
    "300", "", "3.9", "", "existant", 3)
ind("CREC-10", "CREC", "Alertes, ménages ou événements à haut risque ayant reçu un appui CREC",
    "# of high-risk alerts/households/events receiving RCCE support", "Nombre", "Tier 2", "Événements appuyés",
    "N/A", "Rapports de déploiement CREC", "", M, "Type d'événement", "300", "5 000", "", "3.10", "", "existant", 2)
ind("CREC-11", "CREC", "Leaders communautaires, religieux et tradipraticiens engagés",
    "# of community leaders/religious leaders/traditional healers engaged", "Nombre", "Tier 2", "Leaders engagés",
    "N/A", "Rapports d'engagement", "", M, "Type de leader, sexe", "200", "2 500", "", "3.11", "", "existant", 3)
ind("CREC-12", "CREC", "Écoles atteintes par des séances de sensibilisation", "# of schools reached with EVD awareness sessions",
    "Nombre", "Tier 2", "Écoles atteintes", "N/A", SBC, "type_site = École", M, "ZS", "100", "", "", "3.12", "",
    "existant", 3)
ind("CREC-13", "CREC", "Mécanismes de retour communautaire établis", "# of community feedback loops established",
    "Nombre", "Tier 2", "Canaux de retour fonctionnels", "N/A", "Rapports du système de retour", "", M, "Type de canal",
    "15", "150 (milieux de déplacement)", "", "3.13", "", "existant", 3)
ind("CREC-14", "CREC", "Épisodes de résistance communautaire résolus en 48 h",
    "% of community resistance events resolved", "Pourcentage", "Tier 1", "Réticences levées dans les 48 h",
    "Réticences enregistrées", SBC, "reticences_levees_48h (nouveau) / reticences_enregistrees", M, "ZS",
    "≥ 80 %", "", "< 80 % : rouge", "3.14", "", "à corriger", 1, "Ajouter le délai de 48 h et l'aligner sur la fiche SBC.")
ind("CREC-15", "CREC", "Retours communautaires clôturés avec réponse à la communauté",
    "% of feedback items closed with response to the community", "Pourcentage", "Tier 1",
    "Retours clôturés avec réponse communiquée", "Retours enregistrés", "Registre des retours", "", M,
    "Catégorie", "", "", "", "", "", "nouveau", 2)
ind("CREC-16", "CREC", "Typologie hebdomadaire des rumeurs", "Distribution of rumors by category", "Ratio",
    "Contexte", "Rumeurs par catégorie (déni, peur du CTE, EDS, vaccin, stigmatisation...)", "Total des rumeurs",
    "Registre des rumeurs", "", H, "Catégorie, ZS", "", "", "", "", "", "nouveau", 2)
ind("CREC-17", "CREC", "Acceptation des enterrements dignes et sécurisés dans les ZS CGPP",
    "SDB acceptance rate in CGPP areas", "Pourcentage", "Tier 1", "EDS réalisés",
    "Alertes de décès nécessitant un EDS", SIT + " ; " + SBC, "deces_communautaires_eds / deces_communautaires_notifies",
    H, "ZS", "", "", "< 80 % : rouge", "", "", "nouveau", 1,
    "En Ituri, 21 % des EDS échouent (SitRep 149) ; CGPP agit en amont par l'engagement des familles.")
ind("CREC-18", "CREC", "Connaissances clés de la population (enquête rapide)",
    "% of people knowing ≥3 EVD signs and the 151 hotline", "Pourcentage", "Résultat",
    "Personnes interrogées connaissant au moins 3 signes et le numéro vert 151", "Personnes interrogées",
    "Enquête rapide (téléphone ou ménages)", "", Q, "Sexe, âge, ZS", "", "", "", "", "", "nouveau", 3)
ind("CREC-19", "CREC", "Personnel et volontaires CGPP formés à la PSEA (code de conduite signé)",
    "% of CGPP staff and volunteers trained in PSEA", "Pourcentage", "Tier 2",
    "Agents et volontaires formés ayant signé le code de conduite", "Agents et volontaires actifs",
    "Registres RH et de formation", "", M, "IP, sexe", "", "", "< 100 % : rouge", "", "", "nouveau", 1,
    "Le SitRep relève que seules 13 ZS sur 36 en Ituri connaissent le mécanisme de signalement.")

# ----------------------------------------------------------------- PCI / EDS / RÉFÉRENCE
ind("PCI-01", "PCI", "Alertes de décès communautaires référées aux équipes EDS",
    "# of community death alerts referred to SDB teams", "Nombre", "Tier 2", "Alertes de décès référées", "N/A",
    SBC, "deces_communautaires_notifies", H, "ZS", "300", "1 000", "", "4.1", "", "existant", 1)
ind("PCI-02", "PCI", "Événements funéraires éligibles appuyés pour identifier les contacts",
    "% of eligible funeral-linked events supported for contact identification", "Pourcentage", "Tier 1",
    "Événements appuyés", "Événements éligibles", "Journal de surveillance funéraire", "", H, "ZS", "100 %", "",
    "", "4.2", "", "à corriger", 2, "Cible codée « 1 » : écrire 100 %.")
ind("PCI-03", "PCI", "FOSA et sites communautaires appuyés en PCI", "# of health facilities/community sites supported with IPC/readiness inputs",
    "Nombre", "Tier 2", "Sites ayant reçu un appui PCI", "N/A", "Rapports de supervision", "", M, "Type de site, ZS",
    "150", "1 200", "", "4.3", "20", "existant", 2)
ind("PCI-04", "PCI", "Kits EPI/PCI acquis et distribués", "# of PPE/IPC kits or supplies procured and distributed",
    "Nombre", "Tier 2", "Kits distribués", "N/A", "Registres d'achat et de distribution", "", M, "Type de kit",
    "1 500", "250 kits PCI/WASH + 500 kits de préparation", "", "4.4", "19", "à corriger", 2,
    "La cible de 1 500 correspond aux 1 500 agents équipés en EPI (activité 5.6) : ne pas compter deux fois.")
ind("PCI-05", "PCI", "Références (cas suspects, contacts symptomatiques, échantillons) coordonnées et suivies",
    "# of suspected case, symptomatic contact, referral, or sample transport actions coordinated and tracked",
    "Nombre", "Tier 1", "Références coordonnées et suivies", "N/A", "Registre des références ; " + SBC,
    "malades_orientes_ct_cte ; transport d'échantillons", H, "Type (cas suspect, contact, échantillon)", "400",
    "5 000 (milieux de déplacement)", "", "4.5", "18", "existant", 1)
ind("PCI-06", "PCI", "Contacts à haut risque appuyés pendant la quarantaine",
    "# of high-risk contacts supported through quarantine follow-up", "Nombre", "Tier 2", "Contacts appuyés", "N/A",
    "Registres de quarantaine", "", H, "ZS", "1 000", "", "", "4.6", "", "existant", 2)
ind("PCI-07", "PCI", "Séances d'orientation ou de formation PCI", "# of IPC orientation/training sessions conducted",
    "Nombre", "Tier 2", "Séances tenues (et participants formés, à rapporter à part)", "N/A", "Rapports de formation",
    "", M, "Type de participant, sexe", "150", "8 000 participants", "", "4.7", "", "à corriger", 2,
    "Le plan DOS compte des séances, la Tranche 5 des participants : rapporter les deux.")
ind("PCI-08", "PCI", "Circuits de référence cartographiés et mis à jour", "# of referral pathways mapped and updated",
    "Nombre", "Tier 2", "Circuits cartographiés", "N/A", "Cartographies", "", M, "ZS", "20", "", "", "4.8", "",
    "existant", 3)
ind("PCI-09", "PCI", "Références dont l'aboutissement est confirmé", "% of referrals with confirmed completion",
    "Pourcentage", "Tier 1", "Références arrivées à destination", "Références suivies", "Registre des références ; " + SBC,
    "malades_arrives_fosa / malades_orientes_fosa", H, "Type", "≥ 90 %", "", "< 90 % : rouge", "4.9", "", "existant", 1)
ind("PCI-10", "PCI", "Séances d'engagement des familles pour l'acceptation des EDS",
    "# of family engagement sessions conducted for SDB acceptance", "Nombre", "Tier 2", "Séances tenues", "N/A",
    "Rapports d'engagement EDS", "", M, "ZS", "250", "", "", "4.10", "", "existant", 2)
ind("PCI-11", "PCI", "FOSA appuyées avec un score PCI d'au moins 80 %", "% of supported facilities with IPC score ≥80%",
    "Pourcentage", "Résultat", "FOSA appuyées avec score ≥ 80 %", "FOSA appuyées évaluées", "Score card PCI", "", M,
    "ZS", "", "", "", "", "", "nouveau", 2, "Le SitRep montre des scores moyens de 31 à 38 % au Nord-Kivu.")
ind("PCI-12", "PCI", "Infections de prestataires de première ligne dans les ZS CGPP",
    "# of health worker infections in CGPP areas", "Nombre", "Contexte", "PPL confirmés", "N/A", SIT, "", H, "ZS",
    "", "", "Tout cas : alerte", "", "", "nouveau", 1, "Indicateur sentinelle d'échec de la PCI.")
ind("PCI-13", "PCI", "FOSA appuyées avec triage et poste de lavage des mains fonctionnels",
    "% of supported facilities with functional triage and handwashing station", "Pourcentage", "Résultat",
    "FOSA avec triage et lavage des mains fonctionnels", "FOSA appuyées", "Supervision PCI", "", M, "ZS", "", "", "",
    "", "", "nouveau", 2, "Les réunions de coordination signalent le manque d'eau aux PoC et dans les écoles.")
ind("PCI-14", "PCI", "Sites évalués et priorisés en PCI/WASH", "# of IPC/WASH sites assessed", "Nombre", "Tier 2",
    "Sites évalués", "N/A", "Évaluations PCI/WASH", "", M, "Type de site", "", "600", "", "", "", "nouveau", 2,
    "Propre à la Tranche 5.")
ind("PCI-15", "PCI", "Décontaminations (ménages, FOSA, sites) appuyées ou coordonnées",
    "# of decontamination events supported", "Nombre", "Tier 2", "Événements de décontamination", "N/A",
    "Rapports de décontamination", "", H, "Type", "", "800", "", "", "", "nouveau", 2, "Propre à la Tranche 5.")
ind("PCI-16", "PCI", "Visites de supervision ou contrôles ponctuels PCI", "# of IPC supervision/spot-check visits",
    "Nombre", "Tier 2", "Visites réalisées", "N/A", "Rapports de supervision", "", M, "ZS", "", "1 000", "", "", "",
    "nouveau", 2, "Propre à la Tranche 5.")
ind("PCI-17", "PCI", "Agents et acteurs communautaires formés ou recyclés en PCI",
    "# of frontline workers and community actors trained in IPC", "Nombre", "Tier 2", "Participants formés", "N/A",
    "Rapports de formation", "", M, "Sexe, type", "", "8 000", "", "", "", "nouveau", 2, "Propre à la Tranche 5.")

# ----------------------------------------------------------------- LOGISTIQUE
ind("LOG-01", "LOG", "Équipes de suivi des contacts appuyées", "# of contact tracing teams supported", "Nombre",
    "Tier 2", "Équipes appuyées", "N/A", "Registres logistiques", "", M, "ZS", "150", "", "", "5.1", "", "existant", 2)
ind("LOG-02", "LOG", "EIR ou équipes d'investigation appuyées", "# of RRTs/investigation teams supported", "Nombre",
    "Tier 2", "Équipes appuyées", "N/A", "Journal des EIR", "", H, "ZS", "50", "", "", "5.2", "23", "existant", 2)
ind("LOG-03", "LOG", "Smartphones et terminaux de rapportage acquis ou appuyés",
    "# of smartphones/reporting devices procured or supported", "Nombre", "Tier 2", "Terminaux", "N/A",
    "Registres d'achat", "", M, "", "500", "", "", "5.3", "", "existant", 3)
ind("LOG-04", "LOG", "Véhicules, motos et bateaux appuyés", "# of vehicles, motorcycles, and/or boats supported",
    "Nombre", "Tier 2", "Moyens de transport appuyés", "N/A", "Registres de transport", "", M,
    "Type (véhicule, moto, bateau)", "125", "", "", "5.4", "", "à corriger", 3,
    "125 additionne 20 véhicules, 80 motos et 25 déploiements de bateau : désagréger.")
ind("LOG-05", "LOG", "Téléphones, SIM ou forfaits appuyés", "# of phones/SIMs or airtime/data packages supported",
    "Nombre", "Tier 2", "Forfaits appuyés", "N/A", "Registres de distribution", "", M, "", "700", "", "", "5.5", "",
    "existant", 3)
ind("LOG-06", "LOG", "Agents de terrain ayant reçu des EPI", "# of field staff/contact tracers/community teams receiving PPE",
    "Nombre", "Tier 2", "Agents équipés", "N/A", "Registres de distribution", "", M, "IP", "1 500", "", "", "5.6", "",
    "existant", 2)
ind("LOG-07", "LOG", "Déploiements dans les zones difficiles d'accès ou frontalières",
    "# of deployments to hard-to-reach/informal border areas supported", "Nombre", "Tier 2", "Déploiements", "N/A",
    "Journal des déploiements", "", M, "ZS", "100", "100 déploiements de renfort", "", "5.7", "", "existant", 3)
ind("LOG-08", "LOG", "Agents de renfort et superviseurs recrutés", "# of surge staff and supervisors recruited or supported",
    "Nombre", "Tier 2", "Agents recrutés", "N/A", "Registres RH", "", M, "Fonction", "75", "", "", "5.8", "",
    "existant", 3)
ind("LOG-09", "LOG", "Traceurs communautaires sur la liste de réserve", "# of flexible community contact tracers maintained on roster",
    "Nombre", "Tier 2", "Traceurs sur la liste", "N/A", "Liste de réserve", "", M, "ZS", "1 000", "200 listes d'activation",
    "", "5.9", "", "existant", 3)
ind("LOG-10", "LOG", "Visites de supervision de terrain (contacts manqués, refus)", "# of supervisory field visits conducted",
    "Nombre", "Tier 2", "Visites réalisées", "N/A", "Rapports de supervision", "", M, "ZS", "48", "", "", "5.10", "",
    "à corriger", 2, "Même cible (48) que MEL-01 : distinguer supervision opérationnelle et qualité des données.")
ind("LOG-11", "LOG", "Revues des gaps logistiques", "# of logistics gap reviews conducted", "Nombre", "Tier 2",
    "Revues tenues", "N/A", "Rapports logistiques", "", M, "", "35", "", "", "5.11", "", "existant", 3)
ind("LOG-12", "LOG", "Sites de prépositionnement établis", "# of pre-positioned supply locations established",
    "Nombre", "Tier 2", "Sites établis", "N/A", "Rapports de stock", "", M, "", "15", "", "", "5.12", "", "existant", 3)
ind("LOG-13", "LOG", "Briefings sécurité et autorisations de mouvement", "# of security briefings and movement clearances conducted",
    "Nombre", "Tier 2", "Briefings et autorisations", "N/A", "Rapports sécurité", "", M, "", "100", "", "", "5.13", "",
    "existant", 3)
ind("LOG-14", "LOG", "Déploiements couverts par une procédure de sécurité (doublon)",
    "% of field deployments covered by security clearance procedures", "Pourcentage", "Tier 2",
    "Déploiements couverts", "Déploiements", "Journal sécurité", "", M, "", "100 %", "", "", "5.14", "", "à corriger", 3,
    "Doublon de MEL-14 : à supprimer.")
ind("LOG-15", "LOG", "Jours-équipe perdus pour insécurité ou accès", "% of team-days lost to insecurity or access constraints",
    "Pourcentage", "Contexte", "Jours-équipe perdus", "Jours-équipe planifiés", "Journal des déploiements", "", H, "ZS",
    "", "", "", "", "", "nouveau", 2, "Les SitReps et réunions signalent attaques et arrêts d'activités (Mambasa, Biakato).")
ind("LOG-16", "LOG", "Incidents de sécurité affectant les équipes CGPP", "# of security incidents affecting CGPP teams",
    "Nombre", "Contexte", "Incidents", "N/A", "Rapports d'incident", "", H, "ZS, type", "", "", "Tout incident : alerte",
    "", "", "nouveau", 1)
ind("LOG-17", "LOG", "Jours de rupture de stock d'EPI ou d'intrants", "Stock-out days for PPE/IPC supplies", "Nombre",
    "Contexte", "Jours de rupture par site", "N/A", "Fiches de stock", "", M, "Site, article", "", "", "", "", "",
    "nouveau", 3)

# ----------------------------------------------------------------- MEL / CONFORMITÉ
ind("MEL-01", "MEL", "Visites de supervision et de qualité des données", "# of supervision/data quality visits conducted",
    "Nombre", "Tier 2", "Visites DQA réalisées", "N/A", "Rapports de supervision ; grilles DQA", "", M, "Province",
    "48", "", "", "6.1", "24", "existant", 1)
ind("MEL-02", "MEL", "Alertes de suspicion rapportées dans la SIMR (IDSR)", "% of suspected EVD alerts reported through IDSR",
    "Pourcentage", "Tier 1", "Alertes CGPP retrouvées dans la SIMR", "Alertes CGPP détectées", "SIMR/DHIS2 ; " + SBC, "", M,
    "ZS", "100 %", "", "", "6.2", "25", "à corriger", 2,
    "Dans le cadre global, cette ligne porte la définition des visites de supervision : définition à rétablir.")
ind("MEL-03", "MEL", "Exercices de simulation, exercices de préparation et revues après action",
    "# of simulation exercises/readiness drills/AARs conducted", "Nombre", "Tier 2", "Exercices et revues réalisés",
    "N/A", "Rapports SimEx et AAR", "", "Mensuelle ou trimestrielle", "Type", "3",
    "96 simulations + 48 exercices + 40 revues", "", "6.3", "22", "existant", 3)
ind("MEL-04", "MEL", "Mises à jour hebdomadaires transmises au bailleur", "# of weekly donor updates submitted", "Nombre",
    "Tier 2", "Mises à jour transmises", "N/A", "Rapports au bailleur", "", H, "", "16", "", "", "6.4", "", "existant", 1)
ind("MEL-05", "MEL", "Rapports mensuels transmis", "# of monthly reports submitted", "Nombre", "Tier 2", "Rapports",
    "N/A", "Rapports au bailleur", "", M, "", "4", "", "", "6.5", "", "existant", 2)
ind("MEL-06", "MEL", "Notes d'apprentissage publiées", "# of learning briefs published", "Nombre", "Tier 2", "Notes",
    "N/A", "Notes publiées", "", "Continue", "", "2", "≥ 4", "", "6.6", "", "existant", 3)
ind("MEL-07", "MEL", "Indicateurs suivis en temps réel dans le tableau de bord",
    "% of indicators tracked in real-time MEL dashboard", "Pourcentage", "Tier 2", "Indicateurs alimentés automatiquement",
    "Indicateurs de la matrice maîtresse (priorité 1)", "Tableau de bord", "", M, "", "100 %", "100 %", "", "6.7", "",
    "à corriger", 1, "Cible codée « 1 » ; dénominateur à fixer (indicateurs de priorité 1).")
ind("MEL-08", "MEL", "Visites de supervision réalisées selon le plan", "% of supervision visits completed per plan",
    "Pourcentage", "Tier 1", "Visites réalisées", "Visites planifiées", "Plan et rapports de supervision", "", "Continue",
    "Province", "≥ 85 %", "≥ 85 % (rapprochement avec les registres)", "< 85 % : rouge", "6.8", "", "existant", 2)
ind("MEL-09", "MEL", "Rapports soumis à temps", "% of reports submitted on time", "Pourcentage", "Tier 1",
    "Rapports soumis à l'échéance", "Rapports attendus", "Journal des rapports", "", "Continue", "Type de rapport",
    "100 %", "100 %", "< 90 % : rouge", "6.9", "", "à corriger", 1, "Cible codée « 1 ».")
ind("MEL-10", "MEL", "Revues de conformité", "# of compliance reviews conducted", "Nombre", "Tier 2", "Revues", "N/A",
    "Rapports de conformité", "", M, "", "15", "", "", "6.10", "", "existant", 3)
ind("MEL-11", "MEL", "Revues financières", "# of finance reviews conducted", "Nombre", "Tier 2", "Revues", "N/A",
    "Rapports financiers", "", M, "", "15", "", "", "6.11", "", "existant", 3)
ind("MEL-12", "MEL", "Séances de suivi des actions correctives", "# of corrective-action tracking sessions conducted",
    "Nombre", "Tier 2", "Séances", "N/A", "Suivi des actions correctives", "qc_suivi", M, "", "10", "", "", "6.12", "",
    "existant", 3)
ind("MEL-13", "MEL", "Revues des apprentissages avec partenaires et équipes", "# of learning review sessions with partners and field teams",
    "Nombre", "Tier 2", "Séances", "N/A", "Rapports de revue", "", Q, "", "10", "", "", "6.13", "", "existant", 3)
ind("MEL-14", "MEL", "Déploiements couverts par les procédures de sécurité",
    "% of field deployments covered by security clearance procedures", "Pourcentage", "Tier 2", "Déploiements couverts",
    "Déploiements", "Journal sécurité", "", M, "", "100 %", "", "< 100 % : rouge", "6.14", "", "existant", 2)
ind("MEL-15", "MEL", "Complétude et promptitude des rapports hebdomadaires des IP",
    "% of expected weekly IP reports received on time", "Pourcentage", "Tier 1",
    "Rapports IP x ZS reçus le vendredi (dimanche au plus tard)", "Rapports IP x ZS attendus", "Feuille saisie_hebdo",
    "saisie_hebdo", H, "IP, ZS", "", "", "< 80 % : rouge", "", "", "nouveau", 1,
    "Calculé automatiquement par le contrôle qualité : 12,5 % pour les semaines 36 à 41.")
ind("MEL-16", "MEL", "Facteur de vérification des données (DQA)", "Data verification factor (recounted/reported)",
    "Ratio", "Tier 2", "Valeur recomptée dans les registres", "Valeur déclarée", "Visites DQA", "", Q,
    "Indicateur, IP", "", "", "Hors 0,90 à 1,10 : rouge", "", "", "nouveau", 2,
    "À appliquer à 5 indicateurs : SURV-01, SURV-02, SURV-05, CREC-03, PCI-05.")
ind("MEL-17", "MEL", "Anomalies de qualité corrigées en 7 jours ou moins", "% of data quality issues resolved within 7 days",
    "Pourcentage", "Tier 1", "Anomalies corrigées dans les 7 jours", "Anomalies détectées", "Feuille qc_suivi",
    "qc_suivi", H, "Source, gravité", "", "", "< 80 % : rouge", "", "", "nouveau", 1)

# ----------------------------------------------------------------- PRÉPARATION (Tranche 5)
for k, (fr, en, c) in enumerate([
        ("Évaluations de préparation et microplans réalisés", "# of readiness assessments and microplans completed", "150"),
        ("Agents de préparation formés, recyclés et inscrits sur liste d'activation",
         "# of preparedness workforce members trained and rostered", "8 000"),
        ("Listes d'activation provinciales et de ZS tenues à jour", "# of activation rosters maintained", "200"),
        ("Membres d'EIR recyclés", "# of RRT members refreshed", "500"),
        ("Kits de préparation prépositionnés et suivis", "# of preparedness kits pre-positioned", "500"),
        ("Revues d'action précoce et de préparation", "# of early action/readiness reviews", "40")], start=1):
    ind(f"PREP-{k:02d}", "PREP", fr, en, "Nombre", "Tier 2", fr, "N/A", "Rapports de préparation", "", M,
        "Province, ZS", "", c, "", "", "", "nouveau", 2, "Propre à la Tranche 5 (provinces en anneau).")

# ----------------------------------------------------------------- DÉPLACÉS (Tranche 5)
for k, (fr, en, c) in enumerate([
        ("Sites de déplacés et communautés hôtes cartographiés et classés par risque",
         "# of IDP/host-community sites mapped and risk-ranked", "300"),
        ("Points focaux formés et activés en milieu de déplacement", "# of IDP focal persons trained and activated", "1 500"),
        ("Alertes ou préoccupations captées et référées en milieu de déplacement",
         "# of alerts/concerns captured and referred in displacement settings", "5 000"),
        ("Sites de déplacés appuyés en PCI/WASH", "# of IDP sites supported with IPC/WASH inputs", "250"),
        ("Personnes déplacées et hôtes atteintes par des messages adaptés", "# of IDPs and host community members reached", "500 000"),
        ("Références coordonnées et suivies en milieu de déplacement", "# of referrals coordinated in displacement settings", "5 000"),
        ("Mécanismes de retour ou points d'aide établis", "# of feedback channels or help points established", "150")], start=1):
    ind(f"DEPL-{k:02d}", "DEPL", fr, en, "Nombre", "Tier 2", fr, "N/A", "Rapports des sites de déplacés ; " + SBC,
        "type_site = Site de déplacés / Camp de réfugiés", M, "Site, sexe", "", c, "", "", "", "nouveau", 2,
        "Propre à la Tranche 5 (milieux de déplacement).")

ind("CREC-20", "CREC", "Activités de réduction de la stigmatisation et témoignages de guéris",
    "# of stigma-reduction sessions and survivor testimony events", "Nombre", "Tier 2", "Activités", "N/A",
    "Rapports CREC", "", M, "ZS", "", "500", "", "", "", "nouveau", 2, "Propre à la Tranche 5.")
ind("CREC-21", "CREC", "Kits de messages remis aux messagers de confiance", "# of message kits for trusted messengers",
    "Nombre", "Tier 2", "Kits remis", "N/A", "Registres de distribution", "", M, "", "", "2 500", "", "", "",
    "nouveau", 3, "Propre à la Tranche 5.")


# Correspondance : intitulé du classeur hebdomadaire -> (id, désagrégation)
CLASSEUR_VERS_ID = {
    "# of suspected EVD alerts detected through CGPP supported systems": ("SURV-05", ""),
    "# of confirmed Ebola cases identified within CGPP focal areas": ("SURV-17", ""),
    "% of suspected EVD alerts reported by CGPP verified within 24 hours": ("SURV-06", ""),
    "% of suspected EVD alerts notified to authorities within 24 hours": ("SURV-07", ""),
    "% of suspected EVD alerts investigated within 7 days": ("SURV-08", ""),
    "# of CBS actors/community volunteers trained by CGPP on EVD": ("SURV-09", ""),
    "# of risk mapping exercises conducted": ("SURV-10", ""),
    "# of contacts line-listed and monitored": ("SURV-01", ""),
    "% of listed contacts followed daily by CGPP during 21-day follow-up": ("SURV-02", ""),
    "# of households reached with RCCE messages": ("CREC-01", ""),
    "# of targetted RCCE group sessions conducted": ("CREC-02", ""),
    "# of people reached with EVD RCCE messaging house to house visits": ("CREC-03", "canal=VAD"),
    "# of people reached with EVD RCCE messaging during group sessions": ("CREC-03", "canal=séances de groupe"),
    "# of targeted RCCE sessions conducted in high-risk settings": ("CREC-04", ""),
    "# of rumors or community concerns logged": ("CREC-05", ""),
    "% of rumors/community concerns addressed within 48 hours": ("CREC-06", ""),
    "# of people screened at PoEs/PoCs or mobility points": ("SURV-14", ""),
    "# of operational screening points/PoEs established or supported": ("SURV-13", ""),
    "# of suspected case reffered": ("PCI-05", "type=cas suspect"),
    "# of symptomatic contact reffered": ("SURV-19", ""),
    "# of sample transport actions coordinated and tracked": ("PCI-05", "type=transport d'échantillons"),
    "# of PPE/IPC kits or supplies procured and distributed": ("PCI-04", ""),
    "# of health facilities/community sites supported with IPC/readiness inputs": ("PCI-03", ""),
    "# of coordination meetings participated in or supported": ("COORD-03", ""),
    "# of simulation exercises/readiness drills/AARs conducted": ("MEL-03", ""),
    "# of RRT trained and supported by CGPP": ("SURV-12", ""),
    "# of supervision/data quality visits conducted": ("MEL-01", ""),
}


def matrice():
    import pandas as pd
    return pd.DataFrame(I, columns=COLS)


if __name__ == "__main__":
    m = matrice()
    print(m.groupby(["pilier", "statut"]).size().unstack(fill_value=0))
    print("Total :", len(m), "| doublons d'ID :", m.id_indicateur.duplicated().sum())
