const CONFIG = {
  clientId: "b14a3723-db40-4b52-b8b5-e8f504928156",
  tenantId: "4374c25f-cdc7-424b-9931-f5436ed865f1",
  siteId: "assodamesdelaprovidenceorg.sharepoint.com,47544e42-288e-4d35-b949-e6df27098373",
};

const LISTES = {
  Vehicules: "f332f29d-f157-4536-ac6f-b176ac96dd0d",
  Salaries:  "4891187b-3c79-49cb-a8ac-efeaafc84d4f",
  Trajets:   "1b007346-e596-48f9-bf7d-1d28797e3277",
};

const COLS = {
  Vehicule:           "V_x00e9_hicule",
  Conducteur:         "Conducteur",
  Date_Depart:        "DateD_x00e9_part",
  Km_Depart:          "KmD_x00e9_part",
  Motif:              "Motif",
  Commentaire_Depart: "Commentaired_x00e9_part",
  Dommage_Depart:     "DommagesD_x00e9_part",
  Date_Retour:        "DateRetour",
  Km_Retour:          "KmRetour",
  Km_Parcourus:       "KmParcourus",
  Commentaire_Retour: "CommentaireRetour",
  Statut:             "Statut0",
};

const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: "https://cahier-bord-canopee.netlify.app/index.html",
  },
  cache: { cacheLocation: "sessionStorage" },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let msalInitialise = false;
const scopes = ["User.Read", "Sites.ReadWrite.All"];

let etat = {
  utilisateur: null,
  vehiculeId: null,
  vehiculeNom: null,
  vehiculeUnite: null,
  vehiculeKmActuel: 0,
  vehiculeItemId: null,
  vehiculeSharePointId: null,
  trajetEnCoursId: null,
  trajetEnCoursDetail: null,
};

window.addEventListener("load", async () => {
  const params = new URLSearchParams(window.location.search);
  etat.vehiculeId = params.get("v") || null;

  const maintenant = new Date();
  const dateStr = maintenant.toISOString().split("T")[0];
  const heureStr = maintenant.toTimeString().slice(0, 5);
  document.getElementById("date-depart").value = dateStr;
  document.getElementById("heure-depart").value = heureStr;
  document.getElementById("date-retour").value = dateStr;
  document.getElementById("heure-retour").value = heureStr;

  if (!msalInitialise) { await msalInstance.initialize(); msalInitialise = true; }

  try {
    const resultat = await msalInstance.handleRedirectPromise();
    if (resultat) {
      msalInstance.setActiveAccount(resultat.account);
      await initialiserApresConnexion();
      return;
    }
  } catch(e) { console.error("Erreur redirect:", e); }

  const comptes = msalInstance.getAllAccounts();
  if (comptes.length > 0) {
    msalInstance.setActiveAccount(comptes[0]);
    await initialiserApresConnexion();
  }
});

async function seConnecter() {
  try {
    if (!msalInitialise) { await msalInstance.initialize(); msalInitialise = true; }
    const resultat = await msalInstance.loginPopup({ scopes });
    msalInstance.setActiveAccount(resultat.account);
    await initialiserApresConnexion();
  } catch (err) {
    console.error("Erreur connexion:", err);
    alert("Erreur de connexion : " + err.message);
  }
}

function seDeconnecter() {
  msalInstance.logoutPopup();
  allerVers("ecran-connexion");
}

async function obtenirToken() {
  if (!msalInitialise) { await msalInstance.initialize(); msalInitialise = true; }
  const compte = msalInstance.getActiveAccount();
  if (!compte) return null;
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes, account: compte });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ scopes });
    return r.accessToken;
  }
}

async function initialiserApresConnexion() {
  const compte = msalInstance.getActiveAccount();
  if (!compte) return;
  etat.utilisateur = compte;
  document.getElementById("nom-utilisateur").textContent = compte.name || compte.username;
  await Promise.all([recupererVehicule(), recupererSalaries(), verifierTrajetEnCours()]);
  allerVers("ecran-accueil");
}

function allerVers(idEcran) {
  document.querySelectorAll(".ecran").forEach(e => { e.classList.remove("actif"); e.style.display = "none"; });
  const ecran = document.getElementById(idEcran);
  ecran.style.display = "flex";
  ecran.classList.add("actif");
  window.scrollTo(0, 0);
}

async function appelGraph(url, methode = "GET", corps = null) {
  const token = await obtenirToken();
  const options = {
    method: methode,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly",
    },
  };
  if (corps) options.body = JSON.stringify(corps);
  const rep = await fetch(url, options);
  if (!rep.ok) { const e = await rep.text(); throw new Error(`Erreur API: ${rep.status} — ${e}`); }
  if (rep.status === 204) return null;
  return rep.json();
}

function urlListe(key) {
  return `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteId}/lists/${LISTES[key]}`;
}

async function recupererVehicule() {
  if (!etat.vehiculeId) {
    document.getElementById("vehicule-nom").textContent = "Aucun véhicule scanné";
    document.getElementById("vehicule-unite").textContent = "Scannez le QR code du véhicule";
    return;
  }
  try {
    const data = await appelGraph(`${urlListe("Vehicules")}/items?$expand=fields&$top=200`);
    if (data.value) {
      const veh = data.value.find(i => i.fields.Title === etat.vehiculeId);
      if (veh) {
        etat.vehiculeNom = `${veh.fields.Title} — ${veh.fields.Marque || ""}`;
        etat.vehiculeUnite = veh.fields.Unite || "";
        etat.vehiculeKmActuel = veh.fields.Km_actuel || veh.fields.Km_Actuel || 0;
        etat.vehiculeItemId = veh.id;
        etat.vehiculeSharePointId = veh.fields.id || parseInt(veh.id);
        // Récupérer l'ID numérique SharePoint
        const idNum = veh.fields["@odata.etag"] ? null : null;
        // On récupère l'ID via l'URL de l'item
        const itemUrl = `https://assodamesdelaprovidenceorg.sharepoint.com/sites/Canopee/_api/web/lists(guid'${LISTES.Vehicules}')/items?$select=Id,Title&$filter=Title eq '${etat.vehiculeId}'`;
        document.getElementById("vehicule-nom").textContent = etat.vehiculeNom;
        document.getElementById("vehicule-unite").textContent = etat.vehiculeUnite;
        if (etat.vehiculeKmActuel > 0) {
          document.getElementById("km-depart").placeholder = `Dernier km : ${etat.vehiculeKmActuel}`;
        }
      } else {
        document.getElementById("vehicule-nom").textContent = etat.vehiculeId;
        document.getElementById("vehicule-unite").textContent = "Véhicule non trouvé dans la liste";
      }
    }
  } catch (err) {
    console.error("Erreur véhicule:", err);
    document.getElementById("vehicule-nom").textContent = etat.vehiculeId || "Erreur";
    document.getElementById("vehicule-unite").textContent = "Impossible de charger les données";
  }
}

async function recupererSalaries() {
  try {
    const data = await appelGraph(`${urlListe("Salaries")}/items?$expand=fields&$top=300`);
    const select = document.getElementById("select-conducteur");
    select.innerHTML = '<option value="">Sélectionnez un conducteur</option>';
    if (data.value) {
      data.value
        .filter(i => i.fields.Title && i.fields.Actif !== false)
        .sort((a, b) => (a.fields.Title || "").localeCompare(b.fields.Title || ""))
        .forEach(item => {
          const nom = item.fields.Title || "";
          const prenom = item.fields.field_2 || "";
          const nomComplet = prenom ? `${nom} ${prenom}` : nom;
          const opt = document.createElement("option");
          opt.value = item.id;
          opt.dataset.nom = nomComplet;
          opt.textContent = nomComplet;
          select.appendChild(opt);
        });
    }
    const nomConnecte = etat.utilisateur?.name?.split(" ")[0]?.toLowerCase();
    if (nomConnecte) {
      for (let opt of select.options) {
        if (opt.textContent.toLowerCase().includes(nomConnecte)) { opt.selected = true; break; }
      }
    }
  } catch (err) {
    console.error("Erreur salariés:", err);
    document.getElementById("select-conducteur").innerHTML = '<option value="">Impossible de charger les salariés</option>';
  }
}

async function obtenirIdNumerique(listeKey, graphId) {
  // Récupérer l'ID numérique SharePoint à partir de l'ID Graph
  const data = await appelGraph(`${urlListe(listeKey)}/items/${graphId}?$expand=fields&$select=id,fields`);
  // L'ID numérique est dans les fields sous "id" ou dans l'URL
  const url = `https://assodamesdelaprovidenceorg.sharepoint.com/sites/Canopee/_api/web/lists(guid'${LISTES[listeKey]}')/items`;
  return data;
}

async function verifierTrajetEnCours() {
  if (!etat.vehiculeId) return;
  try {
    const data = await appelGraph(`${urlListe("Trajets")}/items?$expand=fields&$top=50`);
    console.log("TOUS les champs du dernier trajet:", JSON.stringify(data.value[data.value.length-1]?.fields));
    if (data.value) {
      console.log("Trajets en cours trouvés:", JSON.stringify(data.value.map(t => t.fields)));
      const trajet = data.value.find(t => {
        const fields = t.fields;
        const statut = fields["Statut0"] || "";
        if (statut !== "En cours") return false;
        const titre = fields["Title"] || "";
        return titre.startsWith(etat.vehiculeId);
      });
      if (trajet) {
        etat.trajetEnCoursId = trajet.id;
        etat.trajetEnCoursDetail = trajet.fields;
        const conducteur = trajet.fields[COLS.Conducteur];
        const nomConducteur = typeof conducteur === 'object' ? conducteur.Title : (conducteur || "—");
        document.getElementById("detail-trajet-en-cours").textContent =
          `${nomConducteur} · Départ : ${formaterDate(trajet.fields[COLS.Date_Depart])}`;
        document.getElementById("trajet-en-cours").style.display = "flex";
        document.getElementById("btn-depart").style.display = "none";
        document.getElementById("btn-retour").style.display = "block";
        document.getElementById("recap-detail").innerHTML =
          `Conducteur : <strong>${nomConducteur}</strong><br>` +
          `Départ : ${formaterDate(trajet.fields[COLS.Date_Depart])}<br>` +
          `Km départ : ${trajet.fields[COLS.Km_Depart] || "—"} km<br>` +
          `Motif : ${trajet.fields[COLS.Motif] || "—"}`;
      }
    }
  } catch (err) { console.error("Erreur trajet en cours:", err); }
}

async function soumettreDepart() {
  const selectConducteur = document.getElementById("select-conducteur");
  const conducteurGraphId = selectConducteur.value;
  const conducteurNom = selectConducteur.options[selectConducteur.selectedIndex]?.dataset.nom || "";
  const dateDepart = document.getElementById("date-depart").value;
  const heureDepart = document.getElementById("heure-depart").value;
  const kmDepart = document.getElementById("km-depart").value;
  const motif = document.getElementById("select-motif").value;
  const passagers = document.getElementById("select-passagers").value;
  const commentaire = document.getElementById("commentaire-depart").value;
  const dommage = document.getElementById("dommage-depart").checked;

  if (!conducteurGraphId) return afficherMessage("msg-depart", "Veuillez sélectionner un conducteur.", "erreur");
  if (!kmDepart) return afficherMessage("msg-depart", "Veuillez saisir le kilométrage de départ.", "erreur");
  if (!motif) return afficherMessage("msg-depart", "Veuillez sélectionner un motif.", "erreur");

  const btn = document.querySelector("#ecran-depart .btn-principal");
  btn.textContent = "Enregistrement..."; btn.disabled = true;

  try {
    // Récupérer les IDs numériques SharePoint
    const conducteurData = await appelGraph(`${urlListe("Salaries")}/items/${conducteurGraphId}?$expand=fields`);
    const conducteurSpId = conducteurData.fields.id;

    let vehiculeSpId = null;
    if (etat.vehiculeId) {
      const vehData = await appelGraph(`${urlListe("Vehicules")}/items?$expand=fields&$top=200`);
      const veh = vehData.value.find(i => i.fields.Title === etat.vehiculeId);
      if (veh) vehiculeSpId = veh.fields.id;
    }

    const fields = {
      Title: `${etat.vehiculeId || "?"} — ${dateDepart}`,
      [COLS.Date_Depart]: `${dateDepart}T${heureDepart}:00`,
      [COLS.Km_Depart]: parseInt(kmDepart),
      [COLS.Motif]: motif,
      [COLS.Commentaire_Depart]: commentaire,
      [COLS.Dommage_Depart]: dommage,
      [COLS.Statut]: "En cours",
    };

    // Ajouter les Lookups si disponibles
    if (conducteurSpId) fields["ConducteurLookupId"] = conducteurSpId;
    if (vehiculeSpId) fields["V_x00e9_hiculeLookupId"] = vehiculeSpId;

    await appelGraph(`${urlListe("Trajets")}/items`, "POST", { fields });

    document.getElementById("confirmation-titre").textContent = "Trajet démarré !";
    document.getElementById("confirmation-detail").textContent =
      `${conducteurNom} · ${etat.vehiculeNom || etat.vehiculeId}\nDépart à ${heureDepart} · ${kmDepart} km\nMotif : ${motif}`;
    allerVers("ecran-confirmation");
    await verifierTrajetEnCours();
  } catch (err) {
    console.error("Erreur départ:", err);
    afficherMessage("msg-depart", "Erreur : " + err.message, "erreur");
  } finally {
    btn.textContent = "Démarrer le trajet"; btn.disabled = false;
  }
}

function calculerDistance() {
  const kmRetour = parseInt(document.getElementById("km-retour").value) || 0;
  const kmDepart = etat.trajetEnCoursDetail?.[COLS.Km_Depart] || 0;
  const distance = kmRetour - kmDepart;
  const bloc = document.getElementById("calcul-distance");
  if (kmRetour > 0 && distance >= 0) {
    document.getElementById("valeur-distance").textContent = distance;
    bloc.style.display = "block";
  } else { bloc.style.display = "none"; }
}

async function soumettreRetour() {
  const dateRetour = document.getElementById("date-retour").value;
  const heureRetour = document.getElementById("heure-retour").value;
  const kmRetour = document.getElementById("km-retour").value;
  const commentaire = document.getElementById("commentaire-retour").value;
  const dommage = document.getElementById("dommage-retour").checked;

  if (!kmRetour) return afficherMessage("msg-retour", "Veuillez saisir le kilométrage de retour.", "erreur");
  const kmDepart = etat.trajetEnCoursDetail?.[COLS.Km_Depart] || 0;
  const kmParcourus = parseInt(kmRetour) - kmDepart;
  if (kmParcourus < 0) return afficherMessage("msg-retour", "Kilométrage de retour inférieur au départ.", "erreur");

  const btn = document.querySelector("#ecran-retour .btn-principal");
  btn.textContent = "Clôture en cours..."; btn.disabled = true;

  try {
    await appelGraph(`${urlListe("Trajets")}/items/${etat.trajetEnCoursId}/fields`, "PATCH", {
      [COLS.Date_Retour]: `${dateRetour}T${heureRetour}:00`,
      [COLS.Km_Retour]: parseInt(kmRetour),
      [COLS.Km_Parcourus]: kmParcourus,
      [COLS.Commentaire_Retour]: commentaire,
      [COLS.Statut]: "Clôturé",
    });

    if (etat.vehiculeItemId) {
      await appelGraph(`${urlListe("Vehicules")}/items/${etat.vehiculeItemId}/fields`, "PATCH", {
        Km_actuel: parseInt(kmRetour)
      });
    }

    document.getElementById("confirmation-titre").textContent = "Trajet clôturé !";
    document.getElementById("confirmation-detail").textContent =
      `${etat.vehiculeNom || etat.vehiculeId}\n${kmParcourus} km parcourus\nRetour à ${heureRetour}`;
    allerVers("ecran-confirmation");

    etat.trajetEnCoursId = null; etat.trajetEnCoursDetail = null;
    document.getElementById("trajet-en-cours").style.display = "none";
    document.getElementById("btn-depart").style.display = "block";
    document.getElementById("btn-retour").style.display = "none";
  } catch (err) {
    console.error("Erreur retour:", err);
    afficherMessage("msg-retour", "Erreur : " + err.message, "erreur");
  } finally {
    btn.textContent = "Clôturer le trajet"; btn.disabled = false;
  }
}

function afficherMessage(id, texte, type) {
  const el = document.getElementById(id);
  el.textContent = texte; el.className = `message ${type}`; el.style.display = "block";
  setTimeout(() => { el.style.display = "none"; }, 6000);
}

function formaterDate(dateStr) {
  if (!dateStr) return "—";
  return new Date(dateStr).toLocaleDateString("fr-FR", {
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit"
  });
}
