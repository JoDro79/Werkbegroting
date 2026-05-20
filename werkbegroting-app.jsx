import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Tariefbibliotheek (ingebakken) ───────────────────────────────────────────
// [postnr, hoofdstuk, paragraaf, omschrijving, eenheid,
//  inschrijfpr, kostpr, staf, projkost, werkterr, bouww, sleufl, grondv,
//  kabelw, bemal, civiel_d, bemal_d, sleufl_d, kabel_d, lever]
const TARIEF_RAW = [["202010","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Aanbrengen en verwijderen gronddepot.","st",7746.17,31513.11,0.0,0.0,8646.44,22866.67,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202101","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Transporteren materialen uit Logistieke Depot naar Materiaaldepot en/of werkloca","uur",129.53,90.0,0.0,0.0,0.0,0.0,0.0,0.0,90.0,0.0,0.0,0.0,0.0,0.0,0.0],["202130","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen Materiaaldepot op het werk.","week",958.98,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202210","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Inrichten en opruimen werkterrein.","st",21645.54,26447.0,0.0,26447.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202220","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen beveiliging materiaaldepot.","st*wk",286.76,112.66,0.0,112.66,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202230","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen beveiliging sleuf.","st*wk",286.76,170.13,0.0,170.13,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202240","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen tijdelijke afrastering van bouwhekken per locatie.","m",9.26,11.87,0.0,0.55,11.32,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202250","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Instandhouden tijdelijke afrastering van bouwhekken per locatie.","m*wk",0.69,1.1,0.0,0.0,1.1,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202260","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen tijdelijke afrastering schapengaas","m",10.57,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202270","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Instandhouden tijdelijke afrastering schapengaas","m*wk",0.21,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202280","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Uitzetten en markeren werkstrook.","m",2.91,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202310","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","AP04 + PFAS in situ ondergrond van het gronddepot.","keer",2294.11,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202320","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","AP04 + PFAS grond in gronddepot <10.000 ton.","keer",2294.11,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202330","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Nul- en eindopname eenmalige kosten","keer",2573.5,1280.0,0.0,0.0,0.0,0.0,0.0,1280.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202340","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Nul- en eindopname tracé.","km",807.45,145.56,0.0,0.0,0.0,0.0,0.0,145.56,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["202410","VOORBEREIDENDE WERKZAAMHEDEN","Tijdelijke voorzieningen","Toepassen boombescherming.","st",70.22,26.47,0.0,0.0,2.94,0.0,0.0,0.0,0.0,0.0,23.53,0.0,0.0,0.0,0.0],["203010","VOORBEREIDENDE WERKZAAMHEDEN","Toepassen (tijdelijke) werkwegen","Toepassen rijplatenbaan intrede- of uittredepunt hdd-boring.","m",20.15,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["203020","VOORBEREIDENDE WERKZAAMHEDEN","Toepassen (tijdelijke) werkwegen","Toepassen rijplaten werklocaties.","m",20.62,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["203030","VOORBEREIDENDE WERKZAAMHEDEN","Toepassen (tijdelijke) werkwegen","Leveren, aanbrengen en verwijderen zandbaan werkweg.","m",35.49,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["203040","VOORBEREIDENDE WERKZAAMHEDEN","Toepassen (tijdelijke) werkwegen","Instandhouden rijplatenbaan.","m",5.33,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["204010","VOORBEREIDENDE WERKZAAMHEDEN","Inventarisatie kabels en leidingen","Graven proefsleuven op verzoek directie UAV.","st",225.83,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["206010","VOORBEREIDENDE WERKZAAMHEDEN","Vegetatie","Verwijderen bossages.","are",224.86,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["206020","VOORBEREIDENDE WERKZAAMHEDEN","Vegetatie","Maaien vegetatie.","are",3.23,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["206030","VOORBEREIDENDE WERKZAAMHEDEN","Vegetatie","Frezen werkstrook.","are",4.48,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["207010","VOORBEREIDENDE WERKZAAMHEDEN","Straatmeubilair","Opnemen en terugplaatsen verkeersbordpaal.","st",125.97,117.2,0.0,0.0,117.2,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["207030","VOORBEREIDENDE WERKZAAMHEDEN","Straatmeubilair","Opnemen en terugplaatsen komportalen.","st",252.52,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["207040","VOORBEREIDENDE WERKZAAMHEDEN","Straatmeubilair","Opnemen en terugplaatsen lichtmasten.","st",746.72,331.36,0.0,0.0,331.36,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210010","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen betonstraatstenen.","m2",28.37,30.28,0.0,0.0,30.28,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210020","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen straatbakstenen.","m2",34.92,21.05,0.0,0.0,21.05,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210030","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen sierbestrating.","m2",42.72,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210040","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen betontegels.","m2",23.47,21.77,0.0,0.0,21.77,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210050","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen grasbetontegels.","m2",34.87,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210060","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen bedrijfsvloerplaten.","m2",17.91,7.1,0.0,0.0,7.1,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210070","VERHARDINGEN","Elementenverhardingen","Opgenomen elementenverharding vervoeren naar en van materiaaldepot.","m2",8.91,0.95,0.0,0.0,0.95,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210210","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen opsluitbanden in zand.","m",23.93,26.05,0.0,0.0,26.05,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210220","VERHARDINGEN","Elementenverhardingen","Opbreken en aanbrengen opsluitbanden in beton.","m",33.08,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210230","VERHARDINGEN","Elementenverhardingen","Opnemen en aanbrengen trottoirbanden in zand.","m",31.12,38.01,0.0,0.0,38.01,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["210240","VERHARDINGEN","Elementenverhardingen","Opbreken en aanbrengen trottoirbanden in beton.","m",42.68,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["213010","VERHARDINGEN","Asfaltverhardingen","Zagen asfaltverharding.","m",7.68,6.35,0.0,0.0,6.35,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["213020","VERHARDINGEN","Asfaltverhardingen","Opbreken asfaltverharding inclusief tijdelijk dichtblokken.","m2",138.66,14.74,0.0,0.0,14.74,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214010","VERHARDINGEN","Funderingen","Verwijderen funderingslaag, hergebruik.","m2",3.72,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214020","VERHARDINGEN","Funderingen","Verwijderen funderingslaag, afvoeren.","m2",12.4,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214030","VERHARDINGEN","Funderingen","Aanbrengen funderingslaag van granulaat.","m2",4.26,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214110","VERHARDINGEN","Funderingen","Leveren menggranulaat 0/31,5.","ton",14.05,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214120","VERHARDINGEN","Funderingen","Leveren hydraulische menggranulaat 0/31,5.","ton",21.06,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214130","VERHARDINGEN","Funderingen","Leveren betongranulaat 0/40.","ton",17.56,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214310","VERHARDINGEN","Funderingen","Leveren zand voor zandbed.","ton",14.48,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214320","VERHARDINGEN","Funderingen","Leveren zand voor ophoging.","ton",14.48,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["214330","VERHARDINGEN","Funderingen","Leveren teelaarde.","ton",24.83,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220010","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf LS kabelcircuit (geen hinder) zonder persoon in sleuf.","m",7.93,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220020","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf LS kabelcircuit (geen hinder) met persoon in sleuf.","m",10.61,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220030","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf LS kabelcircuit (hinder).","m",13.29,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220040","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf LS kabelcircuit (veel hinder).","m",26.91,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220050","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf MS kabelcircuit(s) (geen hinder) zonder persoon in sleuf.","m",8.97,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220060","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf MS kabelcircuit(s) (geen hinder) met persoon in sleuf.","m",11.72,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220070","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf MS kabelcircuit(s) (hinder).","m",13.97,5.98,0.0,0.0,0.0,0.0,0.0,5.98,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220080","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf MS kabelcircuit(s) (veel hinder).","m",29.0,10.06,0.0,0.0,0.0,0.0,0.0,10.06,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220090","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf MS + LS/Telecom kabelcircuit (geen hinder).","m",11.83,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220100","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf MS + LS/Telecom kabelcircuit (hinder).","m",14.63,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["220110","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuits Openbare grond","Realiseren sleuf MS + LS/Telecom kabelcircuit (veel hinder).","m",32.66,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["221010","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuit Niet openbare grond","Realiseren sleuf MS kabelcircuit(s).","m",18.82,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["221020","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuit Niet openbare grond","Realiseren sleuf MS + LS/Telecom kabelcircuit.","m",32.04,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["221030","GRONDWERK SLEUVEN","Realiseren sleuf kabelcircuit Niet openbare grond","Ontgraven en aanbrengen teelaarde buiten sleufprofiel.","m3",2.21,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222010","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 20 cm extra): Openbare grond - Geen hinder.","m",0.33,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222020","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 50 cm extra): Openbare grond - Geen hinder.","m",0.81,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222030","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 20 cm extra): Openbare grond - Hinder.","m",2.56,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222040","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 50 cm extra): Openbare grond - Hinder.","m",5.44,11.34,0.0,0.0,0.0,0.0,0.0,11.34,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222050","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 20 cm extra): Openbare grond - Veel hinder.","m",5.38,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222060","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 50 cm extra): Openbare grond - Veel hinder.","m",12.52,17.18,0.0,0.0,0.0,0.0,0.0,17.18,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222070","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 20 cm extra): Niet openbare grond.","m",0.29,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["222080","GRONDWERK SLEUVEN","Bredere sleuf (extra breedte)","Realiseren extra breedte sleuf (per 50 cm extra): Niet openbare grond.","m",0.74,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["223010","GRONDWERK SLEUVEN","Grotere gronddekking (overdiepte) op kabel/leiding","Realiseren grotere gronddekking op kabel/leiding sleuf (per 10 cm extra).","m",0.33,1.57,0.0,0.0,0.0,0.0,0.0,1.57,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["226010","GRONDWERK SLEUVEN","Grondverbeteringen in sleuf","Leveren en verwerken aanvulzand (Type A).","m3",37.17,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["226020","GRONDWERK SLEUVEN","Grondverbeteringen in sleuf","Leveren en verwerken aanvulzand (Type B).","m3",95.62,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["226030","GRONDWERK SLEUVEN","Grondverbeteringen in sleuf","Leveren en verwerken alternatief aanvulzand (IJsselmeerzand).","m3",34.13,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["226040","GRONDWERK SLEUVEN","Grondverbeteringen in sleuf","Leveren en verwerken aanvulzand (zandbed in sleuf).","m3",34.17,27.78,0.0,0.0,0.0,0.0,0.0,27.78,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["227010","GRONDWERK SLEUVEN","Grond vervoeren","Grond vervoeren naar gronddepot of van gronddepot naar sleuf.","m3",14.95,6.98,0.0,0.0,0.0,0.0,0.0,6.98,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["228010","GRONDWERK SLEUVEN","Grond afvoeren uit depot","Grond afvoeren naar verwerkingslocatie max. 25km.","m3",7.51,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["228020","GRONDWERK SLEUVEN","Grond afvoeren uit depot","Grond afvoeren naar verwerkingslocatie max. 50km.","m3",15.03,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["261010","GROENVOORZIENINGEN","Cultuurtechnisch grondwerk openbare grond","Frezen bermen/werkstrook.","are",4.94,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["261020","GROENVOORZIENINGEN","Cultuurtechnisch grondwerk openbare grond","Egaliseren bermen/werkstrook.","are",5.8,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["262010","GROENVOORZIENINGEN","Inzaaien","Inzaaien van bermen/werkstroken.","are",9.44,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["301010","KRUISINGEN","Mantelbuizen in open ontgraving","Aanbrengen mantelbuizen in sleuf.","m",4.43,8.47,0.0,0.0,0.0,0.0,0.0,0.0,8.47,0.0,0.0,0.0,0.0,0.0,0.0],["301020","KRUISINGEN","Mantelbuizen in open ontgraving","Aanbrengen mantelbuizen wegkruising.","m",4.43,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["302010","KRUISINGEN","Persingen","Aanbrengen, instandhouden en verwijderen perslocatie.","st",2126.33,175.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,175.0,0.0,0.0],["302020","KRUISINGEN","Persingen","Maken pneumatische boring met stalen buis Ø 168mm (DN150).","m",98.45,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["302030","KRUISINGEN","Persingen","Maken pneumatische boring met stalen buis Ø 219mm (DN200).","m",120.71,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["302040","KRUISINGEN","Persingen","Maken pneumatische boring met stalen buis Ø 273mm (DN250).","m",165.78,120.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,120.0,0.0,0.0],["302050","KRUISINGEN","Persingen","Maken pneumatische boring met stalen buis Ø 355mm (DN350).","m",248.72,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["303010","KRUISINGEN","Boogzinkers","Aanbrengen, instandhouden en verwijderen opstelplaats boogzinkers.","st",307.72,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["303020","KRUISINGEN","Boogzinkers","Aanbrengen boogzinker PE100 SDR 11 1xØ75 t/m 110mm.","m",62.59,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["303030","KRUISINGEN","Boogzinkers","Aanbrengen boogzinker PE100 SDR 11 3xØ110mm.","m",250.49,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["303040","KRUISINGEN","Boogzinkers","Aanbrengen boogzinker PE100 SDR 11 1xØ160 mm.","m",135.05,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304010","KRUISINGEN","Gestuurde hdd-boringen","Aanbrengen, instandhouden en verwijderen boorlocatie hdd-boring L < 101 m.","st",1743.65,343.47,0.0,0.0,0.0,20.38,0.0,150.36,0.0,0.0,0.0,0.0,172.73,0.0,0.0],["304020","KRUISINGEN","Gestuurde hdd-boringen","Aanbrengen, instandhouden en verwijderen boorlocatie hdd-boring L = 101 t/m 250m","st",28530.16,647.68,0.0,0.0,0.0,38.42,0.0,283.54,0.0,0.0,0.0,0.0,325.71,0.0,0.0],["304030","KRUISINGEN","Gestuurde hdd-boringen","Aanbrengen, instandhouden en verwijderen boorlocatie hdd-boring L = 251 t/m 600m","st",24947.69,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304040","KRUISINGEN","Gestuurde hdd-boringen","Aanbrengen, instandhouden en verwijderen boorlocatie hdd-boring L = 601 t/m 950m","st",36320.51,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304110","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 1x HDPE Ø110mm.","m",46.66,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304120","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 1x HDPE Ø160mm.","m",96.23,21.35,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,21.35,0.0,0.0],["304130","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 2x HDPE Ø110mm.","m",71.33,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304140","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 2x HDPE Ø160mm.","m",86.48,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304150","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 3x HDPE Ø110mm.","m",99.74,148.95,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,148.95,0.0,0.0],["304160","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 3x HDPE Ø160mm.","m",92.94,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304170","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 6x HDPE Ø110mm.","m",112.04,191.71,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,191.71,0.0,0.0],["304180","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 6x HDPE Ø160mm.","m",151.67,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304190","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 9x HDPE Ø110mm.","m",129.67,38.12,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,38.12,0.0,0.0],["304200","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 9x HDPE Ø160mm.","m",184.61,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304210","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 12x HDPE Ø110mm.","m",129.67,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304220","KRUISINGEN","Gestuurde hdd-boringen","Maken hdd-boring 12x HDPE Ø160mm.","m",184.61,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304310","KRUISINGEN","Gestuurde hdd-boringen","Meerprijs aanbrengen 1 extra HDPE Ø110mm in hdd-boring.","m",7.47,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["304320","KRUISINGEN","Gestuurde hdd-boringen","Meerprijs aanbrengen 1 extra HDPE Ø160mm in hdd-boring.","m",24.2,24.2,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,24.2,0.0,0.0],["308010","KRUISINGEN","Aanvullingen bij hdd-boring","Plaatsbepaling d.m.v. gyroscoopmeting.","dag",5161.75,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["308020","KRUISINGEN","Aanvullingen bij hdd-boring","Toepassen drillgrout.","m",28.68,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["308040","KRUISINGEN","Aanvullingen bij hdd-boring","Aanbrengen kleikist en kwelscherm t.b.v. HDD 1 t/m 3xHDPE.","st",1933.94,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["308050","KRUISINGEN","Aanvullingen bij hdd-boring","Meerprijs aanbrengen kleikist en kwelscherm t.b.v. HDD per extra 3xHDPE.","st",493.24,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["308060","KRUISINGEN","Aanvullingen bij hdd-boring","Toepassen UXOscope.","dag",7398.52,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["309010","KRUISINGEN","Leveren mantelbuizen PVC / PE","Leveren PVC Ø110mm SDR 33 rood.","m",4.31,0.55,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.55],["309020","KRUISINGEN","Leveren mantelbuizen PVC / PE","Leveren PVC Ø160mm SDR 33 rood.","m",7.77,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["309030","KRUISINGEN","Leveren mantelbuizen PVC / PE","Leveren HDPE Ø110mm SDR 17,6 zwart met rode streep.","m",4.85,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["309040","KRUISINGEN","Leveren mantelbuizen PVC / PE","Leveren HDPE Ø110mm SDR 11 zwart met rode streep.","m",6.68,7.4,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,7.4],["309050","KRUISINGEN","Leveren mantelbuizen PVC / PE","Leveren HDPE Ø160mm SDR 17,6 zwart met rode streep.","m",10.09,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["309060","KRUISINGEN","Leveren mantelbuizen PVC / PE","Leveren HDPE Ø160mm SDR 11 zwart met rode streep.","m",14.19,3.83,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,3.83],["340010","KRUISINGEN","Voorbereiding bemaling","Plaatsen en verwijderen van peilbuis.","st",229.41,303.75,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,303.75,0.0,0.0,0.0],["345010","KRUISINGEN","Open bemaling","Aanbrengen en verwijderen open bemaling.","m",4.62,1.27,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,1.27,0.0,0.0,0.0],["345020","KRUISINGEN","Open bemaling","Instandhouden open bemaling.","dag",118.08,302.16,0.0,0.0,0.0,0.0,0.0,0.0,0.0,265.39,0.0,36.77,0.0,0.0,0.0],["345030","KRUISINGEN","Open bemaling","Leveren en aanbrengen drainzand t.b.v. bemaling.","m3",27.55,20.16,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,20.16,0.0,0.0,0.0],["350010","KRUISINGEN","Bronbemaling","Toepassen en instandhouden bronbemaling puntlocatie.","keer",1233.07,516.31,0.0,0.0,0.0,0.0,0.0,0.0,0.0,503.31,0.0,13.0,0.0,0.0,0.0],["350020","KRUISINGEN","Bronbemaling","Toepassen bronbemaling langs sleuf < 250m per locatie.","m",23.18,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["350030","KRUISINGEN","Bronbemaling","Toepassen bronbemaling langs sleuf >250m per locatie.","m",19.44,4.62,0.0,0.0,0.0,0.0,0.0,0.0,0.0,4.54,0.0,0.08,0.0,0.0,0.0],["350040","KRUISINGEN","Bronbemaling","Instandhouden bronbemalingspompen bronbemaling langs sleuf.","st*wk",539.1,274.58,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,274.58,0.0,0.0,0.0],["360010","KRUISINGEN","Bijkomende kosten bemaling","Toepassen ontijzeringsinstallatie 20 m3 voor lozen grondwater.","st",2752.93,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["360020","KRUISINGEN","Bijkomende kosten bemaling","Toepassen ontijzeringsinstallatie 40 m3 voor lozen grondwater.","st",5047.04,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["360030","KRUISINGEN","Bijkomende kosten bemaling","Instandhouden ontijzeringsinstallatie 20 m3.","st*wk",131.92,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["360040","KRUISINGEN","Bijkomende kosten bemaling","Instandhouden ontijzeringsinstallatie 40 m3.","st*wk",260.95,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["400010","KABEL/LEIDINGWERK ELEKTRA","Leggen elektrakabels","Aanbrengen LS kabel 1x4x150mm2 (Al).","m",3.87,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["400020","KABEL/LEIDINGWERK ELEKTRA","Leggen elektrakabels","Aanbrengen MS kabel 1x3x240mm2 (Al).","m",3.94,1.14,0.0,0.0,0.0,0.0,0.0,0.0,1.14,0.0,0.0,0.0,0.0,0.0,0.0],["400030","KABEL/LEIDINGWERK ELEKTRA","Leggen elektrakabels","Aanbrengen MS kabel 1x1x240mm2 (Al).","m",3.94,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["400040","KABEL/LEIDINGWERK ELEKTRA","Leggen elektrakabels","Aanbrengen MS kabel 3x1x240mm2 (Al) gebundeld.","m",11.67,1.06,0.0,0.0,0.0,0.0,0.0,0.0,1.06,0.0,0.0,0.0,0.0,0.0,0.0],["400050","KABEL/LEIDINGWERK ELEKTRA","Leggen elektrakabels","Aanbrengen MS kabel 3x1x630mm2 (Al) gebundeld.","m",11.89,10.51,0.0,0.0,0.0,0.0,0.0,0.0,10.51,0.0,0.0,0.0,0.0,0.0,0.0],["400060","KABEL/LEIDINGWERK ELEKTRA","Leggen elektrakabels","Aanbrengen MS kabel 3x1x800mm2 (Al) gebundeld.","m",14.83,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["400070","KABEL/LEIDINGWERK ELEKTRA","Leggen elektrakabels","Aanbrengen kabel in mantelbuis","m",4.86,2.08,0.0,0.0,0.0,0.0,0.0,0.0,2.08,0.0,0.0,0.0,0.0,0.0,0.0],["410010","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren LS verbindingsmof 4-fase kabel 1x4x150mm2 (Al).","st",283.33,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["410020","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren LS aftakmof.","st",307.41,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["410030","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren LS eindmof.","st",175.79,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["410040","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren MS verbindingsmof 3-fase kabel 1x3x240mm2 (Al).","st",779.99,42.67,0.0,0.0,0.0,0.0,0.0,0.0,42.67,0.0,0.0,0.0,0.0,0.0,0.0],["410050","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren MS overgangmof (10-20kV XLPE 3x240 XLPE 3x1x240).","st",820.14,14.99,0.0,0.0,0.0,0.0,0.0,0.0,14.99,0.0,0.0,0.0,0.0,0.0,0.0],["410060","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren MS verbindingsmof 1-fase kabel 3x1x630mm2 (Al).","set",1043.83,51.2,0.0,0.0,0.0,0.0,0.0,0.0,51.2,0.0,0.0,0.0,0.0,0.0,0.0],["410070","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren MS verbindingsmof 1-fase kabel 3x1x800mm2 (Al).","set",1118.38,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["410080","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren MS eindmof met aarding 3-fase kabel 1x3x240mm2 (Al).","st",461.69,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["410090","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren MS eindmoffen met aarding 1-fase kabel 3x1x630mm2 (Al).","set",509.57,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["410100","MONTAGEWERKZAAMHEDEN ELEKTRA","Monteren moffen","Monteren MS eindmoffen met aarding 1-fase kabel 3x1x800mm2 (Al).","set",518.48,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["412010","MONTAGEWERKZAAMHEDEN ELEKTRA","(De)monteren eindsluitingen","Monteren LS eindsluiting 4x150mm2 (Al).","st",221.96,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["412020","MONTAGEWERKZAAMHEDEN ELEKTRA","(De)monteren eindsluitingen","Monteren MS eindsluiting 3x1x240mm2 (Al).","set",774.27,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["412030","MONTAGEWERKZAAMHEDEN ELEKTRA","(De)monteren eindsluitingen","Monteren MS eindsluiting 3x1x630mm2 (Al).","set",983.32,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["412040","MONTAGEWERKZAAMHEDEN ELEKTRA","(De)monteren eindsluitingen","Monteren MS eindsluiting 3x1x800mm2 (Al).","set",1079.38,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["412050","MONTAGEWERKZAAMHEDEN ELEKTRA","(De)monteren eindsluitingen","Demonteren MS eindsluiting 3x1x240mm2 (Al).","set",355.01,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["412060","MONTAGEWERKZAAMHEDEN ELEKTRA","(De)monteren eindsluitingen","Demonteren MS eindsluiting 3x1x630mm2 (Al).","set",459.41,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["412070","MONTAGEWERKZAAMHEDEN ELEKTRA","(De)monteren eindsluitingen","Demonteren MS eindsluiting 3x1x800mm2 (Al).","set",507.58,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["416010","MONTAGEWERKZAAMHEDEN ELEKTRA","Verwijderen elektrakabels","Kabel verwijderen ongeacht soort en diameter (verschroten).","m",4.83,0.72,0.0,0.0,0.0,0.0,0.0,0.0,0.72,0.0,0.0,0.0,0.0,0.0,0.0],["418010","MONTAGEWERKZAAMHEDEN ELEKTRA","Invoeren kabel/kabelbundel","Invoeren LS kabel in Stations.","st",175.84,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["418020","MONTAGEWERKZAAMHEDEN ELEKTRA","Invoeren kabel/kabelbundel","Invoeren drie stuks 1-fase MS kabels in Stations.","set",527.54,68.57,0.0,0.0,0.0,0.0,0.0,0.0,68.57,0.0,0.0,0.0,0.0,0.0,0.0],["418030","MONTAGEWERKZAAMHEDEN ELEKTRA","Invoeren kabel/kabelbundel","Gat boren in fundering voor kabels/kabelbundel in Stations.","st",68.36,13.39,0.0,0.0,0.0,0.0,0.0,0.0,13.39,0.0,0.0,0.0,0.0,0.0,0.0],["418040","MONTAGEWERKZAAMHEDEN ELEKTRA","Invoeren kabel/kabelbundel","Ingraven kabels als ring","set",234.46,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["420010","AARDING CS - OS (op het veld)","Aarding","Aanbrengen aardelektrode 25 mm² Cu vertind.","m",123.63,1.93,0.0,0.0,0.0,0.0,0.0,0.0,1.93,0.0,0.0,0.0,0.0,0.0,0.0],["420020","AARDING CS - OS (op het veld)","Aarding","Aanbrengen aansluitdraad Cu 25 mm².","m",41.75,3.96,0.0,0.0,0.0,0.0,0.0,0.0,3.96,0.0,0.0,0.0,0.0,0.0,0.0],["441010","KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS","Glasvezel netwerk Alliander","Aanbrengen en testen HDPE glasvezelbuis Ø40mm in sleuf MS-kabel(s).","m",1.74,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["441020","KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS","Glasvezel netwerk Alliander","Aanbrengen en testen HDPE glasvezelbuis Ø40mm in mantel- of kabelbeschermingsbui","m",1.74,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["500010","STATIONS","Stations civiel","Terrein stations ophogen en afwerken.","m3",1.36,7.96,0.0,0.0,0.0,0.0,0.0,7.96,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["500020","STATIONS","Stations civiel","Aanbrengen betontegels rondom stations.","m2",31.48,8.45,0.0,0.0,8.45,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["500030","STATIONS","Stations civiel","Aanbrengen grasbetontegels nabij stations.","m2",25.39,5.82,0.0,0.0,5.82,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["500040","STATIONS","Stations civiel","Coördinatie plaatsen station.","st",279.88,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["800010","BIJKOMENDE WERKZAAMHEDEN","Bijkomende werkzaamheden","Digitaal uitzetten tracé (GPS).","m",0.81,0.37,0.37,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["805010","BIJKOMENDE WERKZAAMHEDEN","Werkplannen","Plan voor het omgaan van vrijgekomen materialen","st",1119.53,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["805020","BIJKOMENDE WERKZAAMHEDEN","Werkplannen","V&G-plan uitvoeringsfase.","st",1119.53,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["806010","BIJKOMENDE WERKZAAMHEDEN","Archeologie","Archeologische begeleiding door senior archeoloog.","dag",963.53,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["807010","BIJKOMENDE WERKZAAMHEDEN","NGE","Benadering verdachte NGE objecten.","dag",3555.87,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["807020","BIJKOMENDE WERKZAAMHEDEN","NGE","Toepassen van een VTVS.","st",2867.64,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["810010","DIRECTIEBEHOEFTEN","Directiebehoeften","Gebruik directieverblijf.","week",1318.28,1737.68,0.0,0.0,1737.68,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["820010","OPLEVERING EN REVISIE","Oplevering","Uitvoeren eindmeting/fingerprintmeting.","circuit",4588.22,1733.33,0.0,0.0,0.0,0.0,0.0,0.0,1733.33,0.0,0.0,0.0,0.0,0.0,0.0],["820020","OPLEVERING EN REVISIE","Oplevering","Opstellen en aanleveren ondergrondse revisiegegevens.","km",2246.88,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["820030","OPLEVERING EN REVISIE","Oplevering","Opstellen en aanleveren bovengrondse revisiegegevens Stations.","st",632.68,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["820040","OPLEVERING EN REVISIE","Oplevering","Opstellen en aanleveren opleverdossier.","st",1119.53,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["870010","TER BESCHIKKING STELLEN","T.b.s. Personeel","T.b.s. verkeersregelaar.","uur",40.15,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["870020","TER BESCHIKKING STELLEN","T.b.s. Personeel","T.b.s. grondwerker.","uur",52.76,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["870030","TER BESCHIKKING STELLEN","T.b.s. Personeel","T.b.s. monteur E.","uur",62.52,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["870040","TER BESCHIKKING STELLEN","T.b.s. Personeel","T.b.s. werkverantwoordelijke E.","uur",102.67,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["870050","TER BESCHIKKING STELLEN","T.b.s. Personeel","T.b.s. koppel stratenmakers.","uur",98.65,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["870060","TER BESCHIKKING STELLEN","T.b.s. Personeel","T.b.s. VP/AVP.","uur",52.76,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["870070","TER BESCHIKKING STELLEN","T.b.s. Personeel","T.b.s. VOP.","uur",61.93,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871010","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. mobiele hydraulische graafmachine.","uur",94.63,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871020","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. minigraver.","uur",75.71,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871030","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. rupskraan (8 ton).","uur",83.16,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871040","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. rupskraan (22 ton).","uur",109.83,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871050","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. wiellader.","uur",108.4,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871060","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. vrachtauto.","uur",78.56,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871070","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. veegzuigwagen.","uur",93.77,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["871080","TER BESCHIKKING STELLEN","T.b.s. Materieel","T.b.s. tractor met grondkar.","uur",77.43,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],["920010","STAARTKOSTEN","Staartkosten","Uitvoeringskosten.","%",20.18,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]];

const TARIEF_MAP = {};
TARIEF_RAW.forEach(r => {
  const aliases = [r[0]];
  if (r[0] === "202101") aliases.push("202120");
  if (r[0] === "202120") aliases.push("202101");
  aliases.forEach(k => { TARIEF_MAP[k] = r; });
});

const CAT_KEYS = ["staf","projkost","werkterr","bouww","sleufl","grondv","kabelw","bemal","civiel_d","bemal_d","sleufl_d","kabel_d","lever"];
const CAT_LABELS = ["Stafkosten","Projectkosten","Werkterrein","Bouwwegen","Sleufloos","Grondverzet","Kabelwerk","Bemaling","Civiel (derden)","Bemaling (derden)","Sleufloos (derden)","Kabelwerk (derden)","Leveranties"];

const euro = v => (v == null || isNaN(v)) ? "—" :
  new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(v);
const pct = v => (v == null || isNaN(v) || !isFinite(v)) ? "—" : (v * 100).toFixed(1) + "%";
const num = v => (v == null || isNaN(v)) ? "—" :
  new Intl.NumberFormat("nl-NL", { maximumFractionDigits: 2 }).format(v);

// ─── XLSX Export ─────────────────────────────────────────────────────────────
function exportToXlsx(projectName, posts, staartkosten) {
  const wb = XLSX.utils.book_new();

  // ── Sheet 1: Projectbegroting ──
  const rows = [
    ["WERKBEGROTING — " + projectName, "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["Postnr","Omschrijving","Eenh.","Hoeveelheid","Inschr.prijs/eenh","Inschrijfsom","Kostprijs/eenh","Interne kosten","Marge (€)","Marge (%)","Stafkosten","Projectkosten","Werkterrein","Bouwwegen","Sleufloos","Grondverzet","Kabelwerk","Bemaling","Civiel (derden)","Bemaling (derden)","Sleufloos (derden)","Kabelwerk (derden)","Leveranties"],
  ];
  let subtotaal_inschrijf = 0, subtotaal_kosten = 0;
  let lastHoofdstuk = "";
  posts.forEach(p => {
    if (p.hoofdstuk !== lastHoofdstuk) {
      rows.push(["", p.hoofdstuk, "","","","","","","","","","","","","","","","","","","","",""]);
      lastHoofdstuk = p.hoofdstuk;
    }
    const inschrijf = p.hoeveelheid * p.inschrijfprijs;
    const kosten = p.hoeveelheid * p.kostprijs;
    const marge = inschrijf - kosten;
    subtotaal_inschrijf += inschrijf;
    subtotaal_kosten += kosten;
    rows.push([
      p.postnr, p.omschrijving, p.eenheid,
      p.hoeveelheid, p.inschrijfprijs, inschrijf, p.kostprijs, kosten,
      marge, inschrijf > 0 ? marge / inschrijf : 0,
      ...CAT_KEYS.map((_,i) => p.hoeveelheid * (p.cats[i] || 0))
    ]);
  });
  rows.push(["","","","","","","","","","","","","","","","","","","","","","",""]);
  rows.push(["SUBTOTAAL UITVOERING","","","","",subtotaal_inschrijf,"",subtotaal_kosten, subtotaal_inschrijf - subtotaal_kosten, subtotaal_inschrijf > 0 ? (subtotaal_inschrijf - subtotaal_kosten)/subtotaal_inschrijf : 0,"","","","","","","","","","","","",""]);

  // Staartkosten
  rows.push(["","","","","","","","","","","","","","","","","","","","","","",""]);
  rows.push(["STAARTKOSTEN","","","","","","","","","","","","","","","","","","","","","",""]);
  let subtotaal_sk_inschrijf = 0, subtotaal_sk_kosten = 0;
  staartkosten.forEach(sk => {
    const bedrag = sk.pct / 100 * subtotaal_inschrijf;
    subtotaal_sk_inschrijf += bedrag;
    subtotaal_sk_kosten += (sk.interneKosten || 0);
    rows.push([sk.postnr, sk.omschrijving, "%", sk.pct, "", bedrag, "", sk.interneKosten || 0, bedrag - (sk.interneKosten || 0), bedrag > 0 ? (bedrag - (sk.interneKosten || 0)) / bedrag : 0,"","","","","","","","","","","","",""]);
  });
  rows.push(["TOTAAL INSCHRIJFSOM","","","","",subtotaal_inschrijf+subtotaal_sk_inschrijf,"",subtotaal_kosten+subtotaal_sk_kosten, (subtotaal_inschrijf+subtotaal_sk_inschrijf)-(subtotaal_kosten+subtotaal_sk_kosten),"","","","","","","","","","","","","",""]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [8,12,50,8,6,14,14,14,14,10,12,12,12,12,12,12,12,12,12,12,12,12,12].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws, "Werkbegroting");

  // ── Sheet 2: Kostensoorten ──
  const catTotals = Array(13).fill(0);
  posts.forEach(p => { p.cats.forEach((c,i) => { catTotals[i] += p.hoeveelheid * c; }); });
  const totalKosten = catTotals.reduce((a,b)=>a+b,0);
  const catRows = [["Kostensoort","Totaal (€)","Aandeel (%)"]];
  CAT_LABELS.forEach((l,i) => { catRows.push([l, catTotals[i], totalKosten > 0 ? catTotals[i]/totalKosten : 0]); });
  catRows.push(["TOTAAL", totalKosten, 1]);
  const ws2 = XLSX.utils.aoa_to_sheet(catRows);
  ws2["!cols"] = [{wch:25},{wch:16},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws2, "Kostensoorten");

  XLSX.writeFile(wb, `Werkbegroting_${projectName.replace(/\s+/g,"_")}.xlsx`);
}

// ─── Parse inschrijfstaat xlsx ────────────────────────────────────────────────
function parseInschrijfstaat(arrayBuffer) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const posts = [];
  const staartkosten = [];
  let projectnaam = "";
  let metaFound = false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // Try to extract project name from filename context - look for 23894 or similar
    const val = row[1] != null ? String(row[1]).trim() : "";
    if (!metaFound && val.length >= 6 && /^\d{6}$/.test(val)) metaFound = true;

    // 6-digit postnr
    if (/^\d{6}$/.test(val)) {
      const hoeveelh = parseFloat(row[4]) || 0;
      const prijs = parseFloat(row[7]) || 0;
      posts.push({ postnr: val, hoeveelheid: hoeveelh, inschrijfprijs: prijs });
    }
    // Staartkosten: 6-digit starting with 9
    if (/^9\d{5}$/.test(val)) {
      const pct = parseFloat(row[7]) || 0;
      const omschr = row[2] ? String(row[2]).trim() : val;
      staartkosten.push({ postnr: val, omschrijving: omschr.split(".")[0], pct, interneKosten: 0 });
    }
  }

  // Try to extract project name from sheet name or first rows
  const sheetMatch = sheetName.match(/\d{4}_\d{4}/) || [];
  projectnaam = sheetMatch[0] || "Project";

  return { posts, staartkosten, projectnaam };
}

// ─── COLOURS ────────────────────────────────────────────────────────────────
const CHAPTER_COLORS = {
  "VOORBEREIDENDE WERKZAAMHEDEN": { bg: "#e8f5e9", accent: "#2e7d32" },
  "VERHARDINGEN":                  { bg: "#e3f2fd", accent: "#1565c0" },
  "GRONDWERK SLEUVEN":             { bg: "#fff3e0", accent: "#e65100" },
  "GROENVOORZIENINGEN":            { bg: "#e8f5e9", accent: "#2e7d32" },
  "KRUISINGEN":                    { bg: "#f3e5f5", accent: "#6a1b9a" },
  "KABEL/LEIDINGWERK ELEKTRA":     { bg: "#fffde7", accent: "#f57f17" },
  "MONTAGEWERKZAAMHEDEN ELEKTRA":  { bg: "#fffde7", accent: "#f57f17" },
  "AARDING CS - OS (op het veld)": { bg: "#fffde7", accent: "#f57f17" },
  "KABEL/LEIDINGWERK INFORMATIE DISTRIBUTIE MS": { bg: "#fffde7", accent: "#f57f17" },
  "STATIONS":                      { bg: "#fce4ec", accent: "#880e4f" },
  "BIJKOMENDE WERKZAAMHEDEN":      { bg: "#eceff1", accent: "#37474f" },
  "DIRECTIEBEHOEFTEN":             { bg: "#eceff1", accent: "#37474f" },
  "OPLEVERING EN REVISIE":         { bg: "#eceff1", accent: "#37474f" },
  "TER BESCHIKKING STELLEN":       { bg: "#eceff1", accent: "#37474f" },
  "STAARTKOSTEN":                  { bg: "#f5f5f5", accent: "#455a64" },
};

// ── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep] = useState("upload"); // upload | review | result
  const [projectName, setProjectName] = useState("");
  const [rawPosts, setRawPosts] = useState([]);
  const [staartkosten, setStaartkosten] = useState([]);
  const [activeTab, setActiveTab] = useState("begroting");
  const [expandedChapters, setExpandedChapters] = useState({});
  const fileRef = useRef();

  const handleFile = useCallback(async (file) => {
    const buf = await file.arrayBuffer();
    const { posts, staartkosten: sk, projectnaam } = parseInschrijfstaat(buf);
    const name = file.name.replace(/\.xlsx$/i,"").replace(/_/g," ");
    setProjectName(name.substring(0, 60));

    // Enrich posts with tarief data
    const enriched = posts
      .filter(p => p.hoeveelheid > 0)
      .map(p => {
        const tar = TARIEF_MAP[p.postnr];
        return {
          postnr: p.postnr,
          omschrijving: tar ? tar[3] : `Post ${p.postnr}`,
          eenheid: tar ? tar[4] : "",
          hoofdstuk: tar ? tar[1] : "OVERIG",
          paragraaf: tar ? tar[2] : "",
          hoeveelheid: p.hoeveelheid,
          inschrijfprijs: p.inschrijfprijs || (tar ? tar[5] : 0),
          kostprijs: tar ? tar[6] : 0,
          cats: tar ? tar.slice(7) : Array(13).fill(0),
          inTarieflib: !!tar,
        };
      });

    setRawPosts(enriched);
    setStaartkosten(sk.length > 0 ? sk : [
      { postnr: "920010", omschrijving: "Uitvoeringskosten", pct: 19.61, interneKosten: 0 },
      { postnr: "930010", omschrijving: "Algemene kosten",   pct: 11.00, interneKosten: 0 },
      { postnr: "940010", omschrijving: "Winst",             pct:  2.00, interneKosten: 0 },
      { postnr: "940020", omschrijving: "Risico",            pct:  2.00, interneKosten: 0 },
    ]);

    // Expand all chapters by default
    const chaps = {};
    enriched.forEach(p => { chaps[p.hoofdstuk] = true; });
    setExpandedChapters(chaps);
    setStep("result");
  }, []);

  const onDrop = useCallback(e => {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  }, [handleFile]);

  // Derived totals
  const totals = useMemo(() => {
    let inschrijf = 0, kosten = 0;
    rawPosts.forEach(p => {
      inschrijf += p.hoeveelheid * p.inschrijfprijs;
      kosten += p.hoeveelheid * p.kostprijs;
    });
    const skInschrijf = staartkosten.reduce((s,sk) => s + sk.pct/100*inschrijf, 0);
    const skKosten = staartkosten.reduce((s,sk) => s + (sk.interneKosten||0), 0);
    return {
      subtotaalInschrijf: inschrijf,
      subtotaalKosten: kosten,
      skInschrijf,
      skKosten,
      totalInschrijf: inschrijf + skInschrijf,
      totalKosten: kosten + skKosten,
      get marge() { return this.totalInschrijf - this.totalKosten; },
      get margePct() { return this.totalInschrijf > 0 ? this.marge / this.totalInschrijf : 0; },
    };
  }, [rawPosts, staartkosten]);

  const catTotals = useMemo(() => {
    const t = Array(13).fill(0);
    rawPosts.forEach(p => p.cats.forEach((c,i) => { t[i] += p.hoeveelheid * c; }));
    return t;
  }, [rawPosts]);

  // Group posts by chapter
  const grouped = useMemo(() => {
    const map = {};
    rawPosts.forEach(p => {
      if (!map[p.hoofdstuk]) map[p.hoofdstuk] = [];
      map[p.hoofdstuk].push(p);
    });
    return map;
  }, [rawPosts]);

  const missingTarief = rawPosts.filter(p => !p.inTarieflib);

  // ─── Render ────────────────────────────────────────────────────────────────
  return (
    <div style={{
      minHeight: "100vh",
      background: "#0f1923",
      fontFamily: "'IBM Plex Mono', 'Courier New', monospace",
      color: "#e8eaed",
    }}>
      {/* Header */}
      <header style={{
        background: "linear-gradient(135deg, #0f1923 0%, #1a2d40 50%, #0f1923 100%)",
        borderBottom: "2px solid #1e4976",
        padding: "18px 28px",
        display: "flex",
        alignItems: "center",
        gap: 16,
        position: "sticky",
        top: 0,
        zIndex: 100,
      }}>
        <div style={{
          width: 38, height: 38,
          background: "linear-gradient(135deg, #1976d2, #0d47a1)",
          borderRadius: 8,
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 18, fontWeight: 900, color: "#fff",
          boxShadow: "0 0 16px rgba(25,118,210,0.5)",
          flexShrink: 0,
        }}>W</div>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, letterSpacing: 3, color: "#90caf9", textTransform: "uppercase" }}>A.HAK ELECTRON</div>
          <div style={{ fontSize: 11, color: "#546e7a", letterSpacing: 2 }}>WERKBEGROTING GENERATOR · HELIX NHN</div>
        </div>
        {step === "result" && (
          <div style={{ marginLeft: "auto", display: "flex", gap: 10, alignItems: "center" }}>
            <button onClick={() => { setStep("upload"); setRawPosts([]); }}
              style={{
                background: "transparent", border: "1px solid #37474f",
                color: "#78909c", padding: "6px 14px", borderRadius: 6,
                cursor: "pointer", fontSize: 11, letterSpacing: 1,
              }}>↩ NIEUW PROJECT</button>
            <button onClick={() => exportToXlsx(projectName, rawPosts, staartkosten)}
              style={{
                background: "linear-gradient(135deg, #1976d2, #0d47a1)",
                border: "none", color: "#fff", padding: "8px 18px",
                borderRadius: 6, cursor: "pointer", fontSize: 11,
                fontWeight: 700, letterSpacing: 1,
                boxShadow: "0 2px 12px rgba(25,118,210,0.4)",
              }}>↓ EXPORTEER XLSX</button>
          </div>
        )}
      </header>

      {/* ─── UPLOAD SCREEN ─── */}
      {step === "upload" && (
        <div style={{ maxWidth: 640, margin: "80px auto", padding: "0 24px" }}>
          <div style={{ textAlign: "center", marginBottom: 48 }}>
            <h1 style={{
              fontSize: 28, fontWeight: 900,
              background: "linear-gradient(90deg, #90caf9, #42a5f5)",
              WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent",
              letterSpacing: 1, margin: "0 0 10px",
            }}>WERKBEGROTING</h1>
            <p style={{ color: "#546e7a", fontSize: 13, letterSpacing: 1 }}>
              LAAD EEN INSCHRIJFSTAAT · ONTVANG EEN WERKBEGROTING
            </p>
          </div>

          <div
            onDrop={onDrop}
            onDragOver={e => e.preventDefault()}
            onClick={() => fileRef.current.click()}
            style={{
              border: "2px dashed #1e4976",
              borderRadius: 16,
              padding: "60px 32px",
              textAlign: "center",
              cursor: "pointer",
              transition: "all 0.2s",
              background: "rgba(25,118,210,0.04)",
              position: "relative",
              overflow: "hidden",
            }}
            onMouseEnter={e => { e.currentTarget.style.borderColor = "#42a5f5"; e.currentTarget.style.background = "rgba(25,118,210,0.08)"; }}
            onMouseLeave={e => { e.currentTarget.style.borderColor = "#1e4976"; e.currentTarget.style.background = "rgba(25,118,210,0.04)"; }}
          >
            <div style={{ fontSize: 48, marginBottom: 16 }}>📋</div>
            <div style={{ fontSize: 15, color: "#90caf9", fontWeight: 700, marginBottom: 8 }}>
              Sleep de inschrijfstaat hierheen
            </div>
            <div style={{ fontSize: 12, color: "#37474f" }}>of klik om een bestand te kiezen · .xlsx</div>
          </div>
          <input ref={fileRef} type="file" accept=".xlsx" style={{ display: "none" }}
            onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />

          <div style={{ marginTop: 32, padding: 20, background: "rgba(255,255,255,0.03)", borderRadius: 10, border: "1px solid #1e4976" }}>
            <div style={{ fontSize: 11, color: "#546e7a", letterSpacing: 1, marginBottom: 10 }}>WAT DOET DEZE TOOL</div>
            {["Leest automatisch alle besteksposten en hoeveelheden",
              "Koppelt aan de HELIX NHN tariefbibliotheek (kostprijzen)",
              "Berekent inschrijfsom vs. interne kosten en marge",
              "Splitst kosten uit per kostensoort",
              "Exporteert direct naar Excel"
            ].map((t,i) => (
              <div key={i} style={{ display: "flex", gap: 10, marginBottom: 6, fontSize: 12, color: "#78909c" }}>
                <span style={{ color: "#1976d2" }}>▸</span> {t}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ─── RESULT SCREEN ─── */}
      {step === "result" && (
        <div style={{ maxWidth: 1400, margin: "0 auto", padding: "20px 20px 60px" }}>
          {/* Project title */}
          <div style={{ marginBottom: 20 }}>
            <input
              value={projectName}
              onChange={e => setProjectName(e.target.value)}
              style={{
                background: "transparent", border: "none", borderBottom: "1px solid #1e4976",
                color: "#90caf9", fontSize: 18, fontWeight: 700, letterSpacing: 2,
                padding: "4px 0", width: "100%", outline: "none",
                textTransform: "uppercase",
              }}
            />
            <div style={{ fontSize: 10, color: "#37474f", marginTop: 4 }}>
              {rawPosts.length} posten geladen · klik op projectnaam om te bewerken
            </div>
          </div>

          {/* KPI bar */}
          <div style={{
            display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 20,
          }}>
            {[
              { label: "INSCHRIJFSOM", value: euro(totals.totalInschrijf), sub: `Excl. staartkost: ${euro(totals.subtotaalInschrijf)}`, color: "#1976d2" },
              { label: "INTERNE KOSTEN", value: euro(totals.totalKosten), sub: `Excl. staartkost: ${euro(totals.subtotaalKosten)}`, color: "#7b1fa2" },
              { label: "BRUTO MARGE", value: euro(totals.marge), sub: pct(totals.margePct) + " van inschrijfsom", color: totals.marge >= 0 ? "#2e7d32" : "#c62828" },
              { label: "POSTEN ACTIEF", value: rawPosts.length, sub: missingTarief.length > 0 ? `⚠ ${missingTarief.length} zonder kostprijs` : "✓ Alle gekoppeld", color: missingTarief.length > 0 ? "#e65100" : "#2e7d32" },
            ].map((k,i) => (
              <div key={i} style={{
                background: "linear-gradient(135deg, #0d1f2d, #132436)",
                border: `1px solid ${k.color}40`,
                borderLeft: `3px solid ${k.color}`,
                borderRadius: 10, padding: "16px 18px",
              }}>
                <div style={{ fontSize: 9, color: "#546e7a", letterSpacing: 2, marginBottom: 6 }}>{k.label}</div>
                <div style={{ fontSize: 20, fontWeight: 900, color: k.color }}>{k.value}</div>
                <div style={{ fontSize: 10, color: "#37474f", marginTop: 4 }}>{k.sub}</div>
              </div>
            ))}
          </div>

          {/* Tab nav */}
          <div style={{ display: "flex", gap: 4, marginBottom: 16, borderBottom: "1px solid #1e4976" }}>
            {[["begroting","BEGROTING"],["staartkosten","STAARTKOSTEN"],["kostensoorten","KOSTENSOORTEN"]].map(([id,lbl]) => (
              <button key={id} onClick={() => setActiveTab(id)} style={{
                background: activeTab===id ? "rgba(25,118,210,0.15)" : "transparent",
                border: "none",
                borderBottom: activeTab===id ? "2px solid #1976d2" : "2px solid transparent",
                color: activeTab===id ? "#90caf9" : "#37474f",
                padding: "8px 18px", cursor: "pointer",
                fontSize: 11, fontWeight: 700, letterSpacing: 2,
                transition: "all 0.15s",
              }}>{lbl}</button>
            ))}
          </div>

          {/* ── TAB: BEGROTING ── */}
          {activeTab === "begroting" && (
            <div>
              {Object.entries(grouped).map(([hoofdstuk, posts]) => {
                const col = CHAPTER_COLORS[hoofdstuk] || { bg: "#1a2d40", accent: "#546e7a" };
                const chInschr = posts.reduce((s,p) => s + p.hoeveelheid*p.inschrijfprijs, 0);
                const chKosten = posts.reduce((s,p) => s + p.hoeveelheid*p.kostprijs, 0);
                const isOpen = expandedChapters[hoofdstuk] !== false;
                return (
                  <div key={hoofdstuk} style={{ marginBottom: 8, borderRadius: 10, overflow: "hidden", border: "1px solid #1e4976" }}>
                    {/* Chapter header */}
                    <div
                      onClick={() => setExpandedChapters(prev => ({ ...prev, [hoofdstuk]: !isOpen }))}
                      style={{
                        background: `${col.accent}18`,
                        borderLeft: `4px solid ${col.accent}`,
                        padding: "10px 16px", cursor: "pointer",
                        display: "flex", alignItems: "center", gap: 12,
                      }}
                    >
                      <span style={{ color: col.accent, fontSize: 12 }}>{isOpen ? "▼" : "▶"}</span>
                      <span style={{ fontWeight: 700, fontSize: 12, letterSpacing: 1, color: col.accent, flex: 1 }}>{hoofdstuk}</span>
                      <span style={{ fontSize: 11, color: "#90caf9" }}>{euro(chInschr)}</span>
                      <span style={{ fontSize: 11, color: "#78909c", marginLeft: 8 }}>intern: {euro(chKosten)}</span>
                      <span style={{
                        fontSize: 11, fontWeight: 700,
                        color: (chInschr-chKosten) >= 0 ? "#4caf50" : "#ef5350",
                        marginLeft: 8,
                      }}>
                        {euro(chInschr-chKosten)}
                      </span>
                    </div>

                    {isOpen && (
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                          <thead>
                            <tr style={{ background: "#0d1520" }}>
                              {["Postnr","Omschrijving","Eenh.","Hoeveelheid","Inschr.prijs","Inschrijfsom","Kostprijs","Int. kosten","Marge €","Marge %"].map(h => (
                                <th key={h} style={{ padding: "6px 10px", color: "#37474f", fontWeight: 700, textAlign: h.includes("som")||h.includes("kosten")||h.includes("prijs")||h.includes("Marge") ? "right" : "left", letterSpacing: 0.5, whiteSpace: "nowrap", borderBottom: "1px solid #1e4976" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {posts.map((p, idx) => {
                              const inschrijf = p.hoeveelheid * p.inschrijfprijs;
                              const kosten = p.hoeveelheid * p.kostprijs;
                              const marge = inschrijf - kosten;
                              const margePct = inschrijf > 0 ? marge/inschrijf : 0;
                              return (
                                <tr key={p.postnr} style={{ background: idx%2===0 ? "transparent" : "rgba(255,255,255,0.02)" }}>
                                  <td style={{ padding: "5px 10px", color: "#546e7a", fontFamily: "monospace" }}>{p.postnr}</td>
                                  <td style={{ padding: "5px 10px", color: p.inTarieflib ? "#cfd8dc" : "#ff8a65", maxWidth: 320 }}>{p.omschrijving}</td>
                                  <td style={{ padding: "5px 10px", color: "#546e7a", textAlign: "center" }}>{p.eenheid}</td>
                                  <td style={{ padding: "5px 10px", textAlign: "right", color: "#90a4ae" }}>{num(p.hoeveelheid)}</td>
                                  <td style={{ padding: "5px 10px", textAlign: "right", color: "#546e7a" }}>{euro(p.inschrijfprijs)}</td>
                                  <td style={{ padding: "5px 10px", textAlign: "right", color: "#90caf9" }}>{euro(inschrijf)}</td>
                                  <td style={{ padding: "5px 10px", textAlign: "right", color: p.kostprijs > 0 ? "#78909c" : "#37474f" }}>{p.kostprijs > 0 ? euro(p.kostprijs) : "—"}</td>
                                  <td style={{ padding: "5px 10px", textAlign: "right", color: p.kostprijs > 0 ? "#ce93d8" : "#37474f" }}>{p.kostprijs > 0 ? euro(kosten) : "—"}</td>
                                  <td style={{ padding: "5px 10px", textAlign: "right", fontWeight: 700, color: marge >= 0 ? "#4caf50" : "#ef5350" }}>{p.kostprijs > 0 ? euro(marge) : "—"}</td>
                                  <td style={{ padding: "5px 10px", textAlign: "right", color: margePct >= 0.1 ? "#4caf50" : margePct >= 0 ? "#ffb74d" : "#ef5350" }}>{p.kostprijs > 0 ? pct(margePct) : "—"}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                );
              })}

              {/* Subtotaal rij */}
              <div style={{
                background: "rgba(25,118,210,0.1)", border: "1px solid #1976d2",
                borderRadius: 8, padding: "12px 20px",
                display: "flex", justifyContent: "space-between", alignItems: "center",
                marginTop: 8,
              }}>
                <span style={{ fontWeight: 700, fontSize: 12, letterSpacing: 2, color: "#90caf9" }}>SUBTOTAAL UITVOERING</span>
                <div style={{ display: "flex", gap: 40 }}>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 10, color: "#546e7a" }}>INSCHRIJFSOM</div>
                    <div style={{ fontSize: 15, color: "#90caf9", fontWeight: 700 }}>{euro(totals.subtotaalInschrijf)}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 10, color: "#546e7a" }}>INTERNE KOSTEN</div>
                    <div style={{ fontSize: 15, color: "#ce93d8", fontWeight: 700 }}>{euro(totals.subtotaalKosten)}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 10, color: "#546e7a" }}>MARGE</div>
                    <div style={{ fontSize: 15, fontWeight: 700, color: totals.subtotaalInschrijf - totals.subtotaalKosten >= 0 ? "#4caf50" : "#ef5350" }}>
                      {euro(totals.subtotaalInschrijf - totals.subtotaalKosten)}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ── TAB: STAARTKOSTEN ── */}
          {activeTab === "staartkosten" && (
            <div style={{ maxWidth: 900 }}>
              <div style={{ marginBottom: 14, fontSize: 12, color: "#546e7a", lineHeight: 1.7 }}>
                De staartkosten worden berekend als percentage over de subtotaalsom uitvoering ({euro(totals.subtotaalInschrijf)}).
                Vul in kolom <strong style={{ color: "#90caf9" }}>Interne kosten</strong> de verwachte eigen inzet in (salariskosten projectleider, uitvoerder, enz.).
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#0d1520", borderBottom: "2px solid #1e4976" }}>
                    {["Post","Omschrijving","% (inschr.)","Inschrijfsom","Interne kosten","Marge"].map(h => (
                      <th key={h} style={{ padding: "8px 14px", color: "#37474f", fontWeight: 700, textAlign: h.includes("Post")||h.includes("Omschr") ? "left" : "right", letterSpacing: 0.5 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {staartkosten.map((sk, i) => {
                    const bedrag = sk.pct / 100 * totals.subtotaalInschrijf;
                    const marge = bedrag - (sk.interneKosten || 0);
                    return (
                      <tr key={sk.postnr} style={{ borderBottom: "1px solid #0d1520", background: i%2===0 ? "transparent" : "rgba(255,255,255,0.02)" }}>
                        <td style={{ padding: "8px 14px", color: "#546e7a", fontFamily: "monospace" }}>{sk.postnr}</td>
                        <td style={{ padding: "8px 14px", color: "#cfd8dc" }}>{sk.omschrijving}</td>
                        <td style={{ padding: "8px 14px", textAlign: "right" }}>
                          <input
                            type="number"
                            value={sk.pct}
                            step={0.01}
                            onChange={e => setStaartkosten(prev => prev.map((s,j) => j===i ? {...s, pct: parseFloat(e.target.value)||0} : s))}
                            style={{
                              background: "rgba(25,118,210,0.1)", border: "1px solid #1e4976",
                              color: "#90caf9", padding: "3px 8px", borderRadius: 4,
                              fontSize: 12, width: 70, textAlign: "right", outline: "none",
                              fontFamily: "inherit",
                            }}
                          />
                          <span style={{ color: "#37474f", marginLeft: 4 }}>%</span>
                        </td>
                        <td style={{ padding: "8px 14px", textAlign: "right", color: "#90caf9" }}>{euro(bedrag)}</td>
                        <td style={{ padding: "8px 14px", textAlign: "right" }}>
                          <input
                            type="number"
                            value={sk.interneKosten || ""}
                            placeholder="0"
                            onChange={e => setStaartkosten(prev => prev.map((s,j) => j===i ? {...s, interneKosten: parseFloat(e.target.value)||0} : s))}
                            style={{
                              background: "rgba(106,27,154,0.1)", border: "1px solid #6a1b9a",
                              color: "#ce93d8", padding: "3px 8px", borderRadius: 4,
                              fontSize: 12, width: 110, textAlign: "right", outline: "none",
                              fontFamily: "inherit",
                            }}
                          />
                        </td>
                        <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 700, color: marge >= 0 ? "#4caf50" : "#ef5350" }}>{euro(marge)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>

              {/* Totaal generaal */}
              <div style={{
                marginTop: 16, background: "rgba(25,118,210,0.1)", border: "1px solid #1976d2",
                borderRadius: 8, padding: "14px 20px",
              }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 20 }}>
                  {[
                    ["TOTAAL INSCHRIJFSOM", euro(totals.totalInschrijf), "#90caf9"],
                    ["TOTAAL INTERNE KOSTEN", euro(totals.totalKosten), "#ce93d8"],
                    ["BRUTO MARGE", `${euro(totals.marge)} · ${pct(totals.margePct)}`, totals.marge >= 0 ? "#4caf50" : "#ef5350"],
                  ].map(([l,v,c]) => (
                    <div key={l}>
                      <div style={{ fontSize: 9, color: "#546e7a", letterSpacing: 2 }}>{l}</div>
                      <div style={{ fontSize: 17, fontWeight: 900, color: c, marginTop: 4 }}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ── TAB: KOSTENSOORTEN ── */}
          {activeTab === "kostensoorten" && (
            <div style={{ maxWidth: 700 }}>
              <div style={{ marginBottom: 16, fontSize: 12, color: "#546e7a" }}>
                Verdeling van de interne kosten over kostensoorten (excl. staartkosten)
              </div>
              {(() => {
                const total = catTotals.reduce((s,v)=>s+v,0);
                return CAT_LABELS.map((lbl, i) => {
                  const val = catTotals[i];
                  const barPct = total > 0 ? val / total : 0;
                  return (
                    <div key={lbl} style={{ marginBottom: 10 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 4 }}>
                        <span style={{ color: "#90a4ae" }}>{lbl}</span>
                        <span style={{ color: val > 0 ? "#ce93d8" : "#37474f" }}>{val > 0 ? `${euro(val)} · ${pct(barPct)}` : "—"}</span>
                      </div>
                      <div style={{ height: 6, background: "#0d1520", borderRadius: 3, overflow: "hidden" }}>
                        <div style={{
                          height: "100%", width: `${barPct * 100}%`,
                          background: `hsl(${200 + i * 15}, 70%, 55%)`,
                          borderRadius: 3, transition: "width 0.6s ease",
                        }} />
                      </div>
                    </div>
                  );
                });
              })()}
              <div style={{
                marginTop: 20, padding: "12px 16px",
                background: "rgba(255,255,255,0.03)", borderRadius: 8,
                border: "1px solid #1e4976",
                display: "flex", justifyContent: "space-between",
              }}>
                <span style={{ fontWeight: 700, fontSize: 12, letterSpacing: 1, color: "#90caf9" }}>TOTAAL INTERNE KOSTEN (UITVOERING)</span>
                <span style={{ fontSize: 14, fontWeight: 900, color: "#ce93d8" }}>{euro(catTotals.reduce((s,v)=>s+v,0))}</span>
              </div>

              {missingTarief.length > 0 && (
                <div style={{ marginTop: 20, padding: 14, background: "rgba(230,81,0,0.08)", border: "1px solid #e65100", borderRadius: 8 }}>
                  <div style={{ fontSize: 11, color: "#ff8a65", fontWeight: 700, marginBottom: 8, letterSpacing: 1 }}>
                    ⚠ {missingTarief.length} POSTEN ZONDER KOSTPRIJS IN TARIEFBIBLIOTHEEK
                  </div>
                  {missingTarief.map(p => (
                    <div key={p.postnr} style={{ fontSize: 11, color: "#78909c", marginBottom: 3 }}>
                      {p.postnr} · {p.omschrijving || "onbekend"} · {euro(p.hoeveelheid * p.inschrijfprijs)}
                    </div>
                  ))}
                  <div style={{ fontSize: 10, color: "#546e7a", marginTop: 8 }}>
                    Inschrijfsom van deze posten telt mee in de totale som, maar interne kosten zijn €0.
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
