#!/usr/bin/env python3
"""
Dashboard GESTION LOGISTIQUE (WMS : commandes clients, préparation, transport, stock et emplacements d'entrepôt).

Lit les tables Excel tb_*.xlsx du dossier Donnees_WMS, calcule les KPI (jointures entre tables),
écrit Export/KPI_Logistique.xlsx (une feuille par jeu de KPI) et génère UN fichier HTML autonome.

  python3 generer_logistique.py                       # dossier Donnees_WMS -> Export/dashboard_logistique.html
  python3 generer_logistique.py --source D:\\WMS --open
  python3 generer_logistique.py --cdn                 # Chart.js en ligne (sinon intégré : téléchargé une fois, gardé dans Python/vendor)
  python3 generer_logistique.py --chartjs chart.umd.js

Tables lues : Orders, Ordres_Line, Order_Preparation, Observations, Truck_Laoding, Truck, Distribution,
Distribution_Line, Regions, Clients, Warehouses, Storage_Bin, Bin_Article, Articles, Restocking, Restocking_Line.
Jamais lues : tb_Users (mots de passe), tb_Audit_Logs. Pas d'analyse fournisseur : la réception n'est suivie que côté entrepôt (Restocking + Restocking_Line, sans Supplier).
"""
from __future__ import annotations

import argparse
import json
import re
import ssl
import sys
import urllib.request
import webbrowser
from pathlib import Path

import pandas as pd


HERE = Path(__file__).resolve().parent
BASE = HERE.parent
SRC = BASE / "Donnees_WMS"
VENDOR = HERE / "vendor" / "chart.umd.js"
URLS = ["https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js", "https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.js",
        "https://unpkg.com/chart.js@4.4.7/dist/chart.umd.js"]
OUT = BASE / "Export" / "dashboard_logistique.html"
XLS = BASE / "Export" / "KPI_Logistique.xlsx"
TEMPLATE = HERE / "dashboard_logistique_template.html"
TABLES = ["Orders", "Ordres_Line", "Order_Preparation", "Observations", "Truck_Laoding", "Truck", "Distribution", "Distribution_Line",
          "Regions", "Clients", "Warehouses", "Storage_Bin", "Bin_Article", "Articles", "Restocking", "Restocking_Line"]
SEUIL_SATURE, SEUIL_VIDE = 0.85, 0.20          # taux d'occupation d'un emplacement


def lire(dossier: Path) -> dict[str, pd.DataFrame]:
    t = {}
    for n in TABLES:
        f = dossier / f"tb_{n}.xlsx"
        if not f.exists():
            sys.exit(f"Table manquante : {f}")
        t[n] = pd.read_excel(f)
    return t


def jours(a, b):                              # a - b en jours (série de dates)
    return (a - b).dt.days


def calculer(t: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    o, pr, ld, di = t["Orders"].copy(), t["Order_Preparation"], t["Truck_Laoding"], t["Distribution"]
    lg = t["Ordres_Line"].groupby("Order_ID").agg(Nb_Lignes=("OrderLine_ID", "count"), Qte=("Quantity", "sum")).reset_index()
    ref = max(o.Order_Date.max(), pr.Preparation_Date.max(), ld.Loading_Date.max(), di.Distribution_Date.max())

    # ---- commandes : une ligne par commande, de la commande à la distribution
    c = (o.merge(t["Clients"][["Customer_ID", "Customer_Name"]], on="Customer_ID", how="left")
          .merge(pr[["Order_ID", "Preparation_ID", "Prepared_By", "Preparation_Date", "Status"]].rename(columns={"Status": "Statut_Prep"}), on="Order_ID", how="left")
          .merge(ld[["Order_ID", "TruckLoading_ID", "Truck_ID", "Loading_Date"]], on="Order_ID", how="left")
          .merge(di[["TruckLoading_ID", "Region_ID", "Distribution_Date", "Destination"]], on="TruckLoading_ID", how="left")
          .merge(t["Truck"][["Truck_ID", "Truck_Code"]], on="Truck_ID", how="left")
          .merge(t["Regions"][["ID_Region", "Nom_Region"]], left_on="Region_ID", right_on="ID_Region", how="left")
          .merge(lg, on="Order_ID", how="left"))
    obs = t["Observations"].merge(pr[["Preparation_ID", "Order_ID"]], on="Preparation_ID", how="left")
    c["Nb_Observations"] = c.Order_ID.map(obs.groupby("Order_ID").size()).fillna(0).astype(int)
    # précision de préparation : observations de type Casse / Qualité / Écart de stock = erreur ; fiabilité stock (proxy) : écart de stock seul
    c["Nb_Obs_Erreur"] = c.Order_ID.map(obs[obs.Category.isin(["Casse", "Qualite", "Ecart de stock"])].groupby("Order_ID").size()).fillna(0).astype(int)
    c["Nb_Obs_Ecart"] = c.Order_ID.map(obs[obs.Category == "Ecart de stock"].groupby("Order_ID").size()).fillna(0).astype(int)
    c["Flag_Precision"] = c.apply(lambda r: None if pd.isna(r.Preparation_ID) else int(r.Nb_Obs_Erreur == 0), axis=1)
    c["Flag_Fiabilite"] = c.apply(lambda r: None if pd.isna(r.Preparation_ID) else int(r.Nb_Obs_Ecart == 0), axis=1)
    # valeur des sorties (prix catalogue x quantité) et In Full (quantités distribuées >= quantités commandées, article par article)
    ol = t["Ordres_Line"].merge(t["Articles"][["Article_ID", "Price"]], on="Article_ID", how="left")
    c["Valeur_Commande"] = c.Order_ID.map((ol.Quantity * ol.Price).groupby(ol.Order_ID).sum()).round(2)
    dl = t["Distribution_Line"].merge(di[["Distribution_ID", "TruckLoading_ID"]], on="Distribution_ID").merge(ld[["TruckLoading_ID", "Order_ID"]], on="TruckLoading_ID")
    qd = dl.groupby(["Order_ID", "Article_ID"]).Quantity.sum().rename("qd").reset_index()
    qc = ol.groupby(["Order_ID", "Article_ID"]).Quantity.sum().rename("qc").reset_index().merge(qd, on=["Order_ID", "Article_ID"], how="left").fillna({"qd": 0})
    full = (qc.qd >= qc.qc).groupby(qc.Order_ID).all()
    c["In_Full"] = c.apply(lambda r: int(full.get(r.Order_ID, False)) if pd.notna(r.Distribution_Date) else None, axis=1)
    c["Delai_Promis_j"] = jours(c.Delivery_Date, c.Order_Date)
    c["Delai_Prep_j"] = jours(c.Preparation_Date, c.Order_Date)
    c["Delai_Charg_j"] = jours(c.Loading_Date, c.Preparation_Date)
    c["Delai_Dist_j"] = jours(c.Distribution_Date, c.Loading_Date)
    c["Cycle_j"] = jours(c.Distribution_Date, c.Order_Date)
    c["Ecart_Prevu_j"] = jours(c.Distribution_Date, c.Delivery_Date)       # > 0 : départ après la date promise
    # OTD : annulée exclue ; distribuée : départ <= date promise ; non distribuée : 0 si la date promise est dépassée
    def otd(r):
        if r.Status == "Annulee":
            return None
        if pd.notna(r.Distribution_Date):
            return int(r.Distribution_Date <= r.Delivery_Date)
        return 0 if r.Delivery_Date < ref else None
    c["Flag_OTD"] = c.apply(otd, axis=1)
    c["Flag_OTIF"] = c.apply(lambda r: None if pd.isna(r.Flag_OTD) else int(r.Flag_OTD == 1 and r.In_Full == 1), axis=1)
    c["Mois"] = c.Order_Date.dt.strftime("%Y-%m")
    c["Annee"] = c.Order_Date.dt.year
    commandes = c.rename(columns={"Order_Code": "Code_Commande", "Customer_Name": "Client", "Order_Date": "Date_Commande", "Delivery_Date": "Date_Livraison_Prevue",
                                  "Status": "Statut", "Prepared_By": "Preparateur", "Preparation_Date": "Date_Preparation", "Loading_Date": "Date_Chargement",
                                  "Distribution_Date": "Date_Distribution", "Nom_Region": "Region", "Truck_Code": "Camion"})[
        ["Order_ID", "Code_Commande", "Client", "Mois", "Annee", "Date_Commande", "Date_Livraison_Prevue", "Statut", "Preparateur", "Statut_Prep", "Date_Preparation",
         "Camion", "Date_Chargement", "Date_Distribution", "Region", "Destination", "Nb_Lignes", "Qte", "Nb_Observations",
         "Delai_Promis_j", "Delai_Prep_j", "Delai_Charg_j", "Delai_Dist_j", "Cycle_j", "Ecart_Prevu_j", "Flag_OTD", "In_Full", "Flag_OTIF",
         "Valeur_Commande", "Nb_Obs_Erreur", "Nb_Obs_Ecart", "Flag_Precision", "Flag_Fiabilite"]]

    # ---- observations de préparation
    ob = obs.merge(c[["Order_ID", "Order_Code"]], on="Order_ID", how="left")
    ob = ob.merge(pr[["Preparation_ID", "Prepared_By"]], on="Preparation_ID", how="left")
    ob["Mois"] = ob.Observation_Date.dt.strftime("%Y-%m")
    ob["Annee"] = ob.Observation_Date.dt.year
    ob["Age_j"] = ((ref - ob.Observation_Date).dt.days).where(ob.Status != "Cloturee")
    observations = ob.rename(columns={"Order_Code": "Code_Commande", "Category": "Categorie", "Description": "Description", "Observation_Date": "Date_Observation",
                                      "Status": "Statut", "Prepared_By": "Preparateur"})[
        ["Observation_ID", "Code_Commande", "Preparateur", "Categorie", "Description", "Date_Observation", "Mois", "Annee", "Statut", "Age_j"]]

    # ---- réceptions entrepôt (dock-to-stock en jours : seules des dates sont fournies, pas d'heures)
    rc = t["Restocking"].merge(t["Warehouses"][["Warehouse_ID", "Warehouse_Name"]], on="Warehouse_ID", how="left") \
        .merge(t["Restocking_Line"].groupby("Restocking_ID").agg(Nb_Lignes=("RestockingLine_ID", "count"), Qte=("Quantity", "sum")).reset_index(), on="Restocking_ID", how="left")
    rc["Delai_Dock_Stock_j"] = jours(rc.Validated_Date, rc.Restock_Date)
    rc["Age_Attente_j"] = (ref - rc.Restock_Date).dt.days.where(rc.Status != "Valide")
    rc["Flag_Valide"] = (rc.Status == "Valide").astype(int)
    rc["Mois"] = rc.Restock_Date.dt.strftime("%Y-%m")
    rc["Annee"] = rc.Restock_Date.dt.year
    receptions = rc.rename(columns={"Restocking_Code": "Code_Reception", "Warehouse_Name": "Entrepot", "Restock_Date": "Date_Reception", "Status": "Statut", "Validated_Date": "Date_Mise_en_Stock"})[
        ["Restocking_ID", "Code_Reception", "Entrepot", "Mois", "Annee", "Date_Reception", "Statut", "Date_Mise_en_Stock", "Nb_Lignes", "Qte", "Delai_Dock_Stock_j", "Age_Attente_j", "Flag_Valide"]]

    # ---- emplacements (photo) et lignes de stock par emplacement
    b = t["Storage_Bin"].merge(t["Warehouses"][["Warehouse_ID", "Warehouse_Name"]], on="Warehouse_ID", how="left")
    b["Taux_Occupation"] = (b.Current_Stock / b.Capacity).round(4)
    b["Zone"] = b.Description.str.replace(r"\s*\(.*\)", "", regex=True).str.strip()
    b["Statut_Emplacement"] = b.Taux_Occupation.apply(lambda x: "Sature" if x >= SEUIL_SATURE else ("Sous-utilise" if x < SEUIL_VIDE else "Normal"))
    emplacements = b.rename(columns={"Warehouse_Name": "Entrepot", "Bin_Code": "Code_Emplacement", "Capacity": "Capacite", "Current_Stock": "Stock_Actuel"})[
        ["Bin_ID", "Code_Emplacement", "Entrepot", "Zone", "Capacite", "Stock_Actuel", "Taux_Occupation", "Statut_Emplacement", "Putaway"]]
    s = (t["Bin_Article"].merge(b[["Bin_ID", "Bin_Code", "Warehouse_Name"]], on="Bin_ID", how="left")
         .merge(t["Articles"][["Article_ID", "Article_Code", "Article_Name", "Price"]], on="Article_ID", how="left"))
    s["Valeur"] = (s.Quantity * s.Price).round(2)
    # Rupture : quantité 0 alors qu'un seuil mini existe ; Vide : quantité 0 sans seuil (emplacement non géré) ; Sous min : 0 < quantité < Min_Level
    def statut(r):
        if r.Quantity == 0:
            return "Rupture" if r.Min_Level > 0 else "Vide"
        return "Sous min" if r.Quantity < r.Min_Level else ("Surstock" if r.Quantity > r.Max_Level else "OK")
    s["Statut_Stock"] = s.apply(statut, axis=1)
    s["Manque_Min"] = (s.Min_Level - s.Quantity).clip(lower=0)
    stock = s.rename(columns={"Warehouse_Name": "Entrepot", "Bin_Code": "Code_Emplacement", "Article_Name": "Article", "Quantity": "Quantite", "Min_Level": "Min", "Max_Level": "Max", "Note": "Note"})[
        ["Bin_Article_ID", "Code_Emplacement", "Entrepot", "Article_Code", "Article", "Quantite", "Min", "Max", "Statut_Stock", "Manque_Min", "Valeur", "Note"]]
    return {"commandes": commandes, "observations": observations, "receptions": receptions, "emplacements": emplacements, "stock": stock,
            "_ref": pd.DataFrame({"ref": [ref.strftime("%Y-%m-%d")]}), "_articles": t["Articles"]}


def en_records(df: pd.DataFrame) -> list[dict]:
    d = df.copy()
    for c in d.columns:
        if pd.api.types.is_datetime64_any_dtype(d[c]):
            d[c] = d[c].dt.strftime("%Y-%m-%d")
    d = d.round(4).astype(object).where(d.notna(), None)
    out = []
    for r in d.to_dict("records"):
        out.append({k: (v.item() if hasattr(v, "item") else v) for k, v in r.items()})
    for r in out:
        for k, v in r.items():
            if isinstance(v, float) and v.is_integer():
                r[k] = int(v)
    return out


def payload(k: dict[str, pd.DataFrame]) -> dict:
    art = k["_articles"].copy()
    art["Valeur"] = art.Price * art.Stock_Quantity
    top = art.sort_values("Valeur", ascending=False).head(10)[["Article_Code", "Article_Name", "Unit", "Price", "Stock_Quantity", "Valeur"]]
    s = k["stock"]
    crit = s[s.Statut_Stock.isin(["Rupture", "Sous min"])].sort_values(["Valeur"], ascending=False).head(400)   # liste de travail : les plus chers d'abord
    agg = s.groupby(["Entrepot", "Statut_Stock"]).agg(n=("Valeur", "size"), valeur=("Valeur", "sum")).reset_index()
    return {"commandes": en_records(k["commandes"]), "observations": en_records(k["observations"]), "receptions": en_records(k["receptions"]),
            "emplacements": en_records(k["emplacements"]), "stock_agg": en_records(agg), "stock_crit": en_records(crit),
            "articles": {"valeur_totale": round(float(art.Valeur.sum()), 2), "nb": int(len(art)), "ruptures": int((art.Stock_Quantity == 0).sum()), "top": en_records(top)},
            "ref": k["_ref"].ref.iloc[0]}


def ecrire_excel(k: dict[str, pd.DataFrame]) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    XLS.parent.mkdir(parents=True, exist_ok=True)
    noms = {"commandes": "KPI_Commandes", "observations": "KPI_Observations", "receptions": "KPI_Receptions", "emplacements": "KPI_Emplacements", "stock": "KPI_Stock_Emplacement"}
    with pd.ExcelWriter(XLS, engine="openpyxl", datetime_format="DD/MM/YYYY", date_format="DD/MM/YYYY") as w:
        for key, name in noms.items():
            df = k[key]
            df.to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            for i, col in enumerate(df.columns, 1):
                h = ws.cell(1, i); h.fill = PatternFill("solid", fgColor="0F2A4A"); h.font = Font(bold=True, color="FFFFFF"); h.alignment = Alignment(horizontal="center")
                ws.column_dimensions[get_column_letter(i)].width = min(34, max(12, len(str(col)) + 3))
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions


def resume(k: dict[str, pd.DataFrame]) -> None:
    c = k["commandes"]; o = c.Flag_OTD.dropna()
    print(f"  Commandes : {len(c)} | livrées {100 * (c.Statut == 'Livree').mean():.1f} % | annulées {100 * (c.Statut == 'Annulee').mean():.1f} % | OTD {100 * o.mean():.1f} % | OTIF {100 * c.Flag_OTIF.dropna().mean():.1f} % | précision prépa {100 * c.Flag_Precision.dropna().mean():.1f} % | "
          f"cycle moyen {c.Cycle_j.mean():.2f} j | préparation {c.Delai_Prep_j.mean():.2f} j")
    ob = k["observations"]; print(f"  Observations : {len(ob)} dont {(ob.Statut != 'Cloturee').sum()} non clôturées")
    rc = k["receptions"]; print(f"  Réceptions : {len(rc)} | validées {100 * rc.Flag_Valide.mean():.1f} % | dock-to-stock {rc.Delai_Dock_Stock_j.mean():.2f} j | en attente {(rc.Flag_Valide == 0).sum()}")
    e = k["emplacements"]; print(f"  Emplacements : occupation moy. {100 * e.Taux_Occupation.mean():.1f} % | saturés {(e.Statut_Emplacement == 'Sature').sum()} | sous-utilisés {(e.Statut_Emplacement == 'Sous-utilise').sum()}")
    s = k["stock"]; print(f"  Stock/emplacement : {len(s)} lignes | rupture {(s.Statut_Stock == 'Rupture').sum()} | sous min {(s.Statut_Stock == 'Sous min').sum()} | surstock {(s.Statut_Stock == 'Surstock').sum()}")


def valide(js: str) -> bool:
    return len(js) > 100_000 and "Chart.js" in js[:400_000]


def charger_chartjs(chemin: Path | None) -> str:
    """Code de Chart.js : fichier fourni, cache local (Python/vendor), ou téléchargement."""
    if chemin:
        js = chemin.read_text(encoding="utf-8")
        if not valide(js):
            sys.exit(f"{chemin} ne ressemble pas à Chart.js (fichier chart.umd.js attendu).")
        return js
    if VENDOR.exists() and valide(VENDOR.read_text(encoding="utf-8")):
        return VENDOR.read_text(encoding="utf-8")
    derniere = None
    for url in URLS:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30, context=ssl.create_default_context()) as r:
                js = r.read().decode("utf-8")
            if valide(js):
                VENDOR.parent.mkdir(exist_ok=True)
                VENDOR.write_text(js, encoding="utf-8")
                return js
        except Exception as e:
            derniere = e
    sys.exit("Téléchargement de Chart.js impossible (" + str(derniere) + f").\nTéléchargez {URLS[0]} puis relancez avec --chartjs CHEMIN, ou utilisez --cdn.")


def generer(source: Path, sortie: Path, chartjs: Path | None, cdn: bool, exemple: bool) -> None:
    k = calculer(lire(source))
    ecrire_excel(k)
    data = json.dumps(payload(k) | {"source": source.name, "sample": exemple}, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8")
    lib = f'<script src="{URLS[0]}"></script>' if cdn else "<script>/* Chart.js (intégré) */\n" + charger_chartjs(chartjs).replace("</script", "<\\/script") + "\n</script>"
    html = html.replace("<!--CHARTJS-->", lib, 1)
    html, n = re.subn(r'(<script id="sample" type="application/json">).*?(</script>)', lambda m: m.group(1) + data + m.group(2), html, count=1, flags=re.S)
    if n != 1:
        sys.exit('Gabarit invalide : balise <script id="sample"> introuvable.')
    sortie.parent.mkdir(parents=True, exist_ok=True)
    sortie.write_text(html, encoding="utf-8")
    print(f"{sortie}  [{'Chart.js en ligne' if cdn else 'autonome'}, {sortie.stat().st_size // 1024} Ko] | {XLS.name} écrit")
    resume(k)


def main() -> None:
    ap = argparse.ArgumentParser(description="Dashboard gestion logistique (WMS).")
    ap.add_argument("--source", type=Path, default=SRC, help="dossier contenant les tb_*.xlsx")
    ap.add_argument("--out", type=Path, default=OUT)
    ap.add_argument("--chartjs", type=Path)
    ap.add_argument("--cdn", action="store_true")
    ap.add_argument("--sample", action="store_true", help="marquer les données comme « exemple »")
    ap.add_argument("--open", action="store_true")
    a = ap.parse_args()
    generer(a.source, a.out, a.chartjs, a.cdn, a.sample)
    if a.open:
        webbrowser.open(a.out.resolve().as_uri())


if __name__ == "__main__":
    main()
