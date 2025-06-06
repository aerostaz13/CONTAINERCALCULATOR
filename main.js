// main.js

let produits = [];      // Contenu de produits.json
let conteneurs = [];    // Contenu de conteneurs.json

// Pour stocker le “dernier résultat” (utilisé lors de l’export Excel)
let lastResult = {
  ref: null,          // { totalVol, totalPds, resultat }
  dry: null,          // id. pour la partie sèche restante
  rowsProduits: []    // [ { Référence, Nom, QtéUnité, QtéParCarton, PoidsCarton, VolCarton, FullCartons, CartonsÀExpédier, VolumeTotal, PoidsTotal, Refrigerer }, … ]
};

/**
 * Au chargement de la page :
 * 1. Charger produits.json et conteneurs.json
 * 2. Générer le tableau des produits (avec inputs pour quantités/carton/poids/vol/carton/réfrigéré)
 * 3. Générer le tableau des conteneurs (modifiable)
 * 4. Brancher les événements des boutons
 */
window.addEventListener("DOMContentLoaded", async () => {
  try {
    const [respP, respC] = await Promise.all([
      fetch("produits.json"),
      fetch("conteneurs.json")
    ]);
    if (!respP.ok || !respC.ok) {
      throw new Error("Impossible de charger les fichiers JSON.");
    }
    produits = await respP.json();
    conteneurs = await respC.json();

    // 2) Tableau des produits
    genererTableProduits();

    // 3) Tableau des conteneurs
    genererTableConteneursOverrides();

    // 4) Boutons
    document.getElementById("btn-calculer")
            .addEventListener("click", traiterCalcul);
    document.getElementById("btn-reset")
            .addEventListener("click", resetForm);
    document.getElementById("btn-download")
            .addEventListener("click", downloadExcel);
  } catch (err) {
    alert("Erreur au chargement des données : " + err.message);
    console.error(err);
  }
});

/**
 * Génère le <tbody> du tableau produits :
 * – Colonne Référence (texte)
 * – Colonne Nom (texte)
 * – Colonne Quantité unitaire (input number)
 * – Colonne Quantité par carton (input number, modifiable)
 * – Colonne Poids Brut Carton (input number, modifiable)
 * – Colonne Volume par Carton (input number, modifiable)
 * – Colonne Réfrigéré ? (checkbox modifiable)
 */
function genererTableProduits() {
  const tbody = document.querySelector("#table-produits tbody");
  tbody.innerHTML = "";

  produits.forEach((prod, i) => {
    const codeRef         = prod["Référence"] || prod["Product"] || "";
    const nomProd         = prod["Product"] + " "+ prod["Presentation"] + " " + prod["Unnamed: 3"] || " ";
    const qtParCartonDef  = parseFloat(prod["Quantité par carton"])   || 1;
    const poidsCartonDef  = parseFloat(prod["Poids Brut Carton"])    || 0;
    const volCartonDef    = parseFloat(prod["M3 par carton"])        || 0;
    const isRefrig        = prod["Refrigerer"] == 1;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${codeRef}</td>
      <td>${nomProd}</td>
      <td>
        <input
          type="number"
          id="quantite-${i}"
          min="0"
          step="1"
          value="0"
          style="width: 60px;"
        />
      </td>
      <td>
        <input
          type="number"
          id="prod-qtparcarton-${i}"
          min="1"
          step="1"
          value="${qtParCartonDef}"
          style="width: 60px;"
        />
      </td>
      <td>
        <input
          type="number"
          id="prod-poidscarton-${i}"
          min="0"
          step="0.001"
          value="${poidsCartonDef.toLocaleString("en", { useGrouping: false, minimumFractionDigits: 3 })}"
        />
      </td>
      <td>
        <input
          type="number"
          id="prod-volcarton-${i}"
          min="0"
          step="0.000001"
          value="${volCartonDef.toLocaleString("en", { useGrouping: false, minimumFractionDigits: 6 })}"
        />
      </td>
      <td style="text-align: center;">
        <input
          type="checkbox"
          id="prod-refrig-${i}"
          ${isRefrig ? "checked" : ""}
        />
      </td>
    `;
    tbody.appendChild(tr);
  });
}

/**
 * Génère le <tbody> du tableau “Capacités des conteneurs (modifiable)” :
 * – Colonne Code conteneur
 * – Colonne Capacité volume (input number)
 * – Colonne Capacité poids (input number)
 */
function genererTableConteneursOverrides() {
  const tbody = document.querySelector("#table-conteneurs tbody");
  tbody.innerHTML = "";

  conteneurs.forEach((cont, i) => {
    const codeCont = (cont["NAME "] || "").trim();
    const volDef   = parseFloat(cont["Capacite_plus_de_quatre"]);
    const pdsDef   = parseFloat(cont["Poids_max"]);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${codeCont}</td>
      <td>
        <input
          type="number"
          id="cont-vol-${i}"
          min="0"
          step="0.000001"
          value="${volDef.toLocaleString("en", { useGrouping: false, minimumFractionDigits: 6 })}"
        />
      </td>
      <td>
        <input
          type="number"
          id="cont-pds-${i}"
          min="0"
          step="0.001"
          value="${pdsDef.toLocaleString("en", { useGrouping: false, minimumFractionDigits: 3 })}"
        />
      </td>
    `;
    tbody.appendChild(tr);
  });
}

/**
 * Met à jour en mémoire :
 *  - prod["Quantité par carton"]
 *  - prod["Poids Brut Carton"]
 *  - prod["M3 par carton"]
 *  - prod["Refrigerer"]
 * d’après les inputs du tableau produits
 */
function updateProduitsFromOverrides() {
  produits.forEach((prod, i) => {
    // Quantité par carton
    const qtCartonInput = document.getElementById(`prod-qtparcarton-${i}`);
    if (qtCartonInput) {
      const newQtCarton = parseInt(qtCartonInput.value, 10);
      if (!isNaN(newQtCarton) && newQtCarton > 0) {
        prod["Quantité par carton"] = newQtCarton;
      }
    }
    // Poids Brut Carton
    const poidsCartonInput = document.getElementById(`prod-poidscarton-${i}`);
    if (poidsCartonInput) {
      const newPoidsCarton = parseFloat(poidsCartonInput.value);
      if (!isNaN(newPoidsCarton)) {
        prod["Poids Brut Carton"] = newPoidsCarton;
      }
    }
    // Volume par Carton
    const volCartonInput = document.getElementById(`prod-volcarton-${i}`);
    if (volCartonInput) {
      const newVolCarton = parseFloat(volCartonInput.value);
      if (!isNaN(newVolCarton)) {
        prod["M3 par carton"] = newVolCarton;
      }
    }
    // Réfrigéré ?
    const checkbox = document.getElementById(`prod-refrig-${i}`);
    if (checkbox) {
      prod["Refrigerer"] = checkbox.checked ? 1 : 0;
    }
  });
}

/**
 * Met à jour en mémoire les capacités volume/poids des conteneurs,
 * d’après ce que l’utilisateur a saisi dans les inputs du tableau conteneurs.
 */
function updateConteneursFromOverrides() {
  conteneurs.forEach((cont, i) => {
    const volInput = document.getElementById(`cont-vol-${i}`);
    const pdsInput = document.getElementById(`cont-pds-${i}`);
    if (volInput && pdsInput) {
      const newVol = parseFloat(volInput.value);
      const newPds = parseFloat(pdsInput.value);
      if (!isNaN(newVol)) {
        cont["Capacite_plus_de_quatre"] = newVol;
      }
      if (!isNaN(newPds)) {
        cont["Poids_max"] = newPds;
      }
    }
  });
}

/**
 * Lorsqu’on clique sur “Calculer” :
 * 1) On récupère toutes les modifications (quantités, carton, poids, volume, réfrigéré).
 * 2) Pour chaque produit, on calcule :
 *    - nb de cartons entiers
 *    - cartons à expédier (entier supérieur si reste)
 *    - poids total (cartons * poids brut carton)
 *    - volume total (cartons * volume/carton)
 *    - message si carton incomplet
 * 3) On agrège en “réfrigéré” vs “non réfrigéré”
 * 4) On trouve les conteneurs optimaux (réfrigérés puis secs)
 * 5) On affiche le résultat (+ messages sur cartons incomplets)
 */
function traiterCalcul() {
  // 1) Mettre à jour en mémoire toutes les infos produit & conteneur
  updateProduitsFromOverrides();
  updateConteneursFromOverrides();

  // Initialisation des totaux
  let totalRefVol = 0, totalRefPds = 0;
  let totalDryVol = 0, totalDryPds = 0;
  lastResult.rowsProduits = [];
  const missingMessages = [];

  // 2) Pour chaque produit, calculer cartons et totaux
  produits.forEach((prod, i) => {
    const qtUnits         = parseInt(document.getElementById(`quantite-${i}`).value, 10) || 0;
    const isRefrig        = prod["Refrigerer"] == 1;
    const unitsPerCarton  = parseInt(prod["Quantité par carton"], 10) || 1;
    const poidsParCarton  = parseFloat(prod["Poids Brut Carton"])    || 0;
    const volParCarton    = parseFloat(prod["M3 par carton"])        || 0;

    if (qtUnits <= 0) {
      // Aucune unité ⇒ on stocke row vide pour l’export
      lastResult.rowsProduits.push({
        Référence:        prod["Référence"],
        Nom:              prod["Product"] + " "+ prod["Presentation"] + " " + prod["Unnamed: 3"],
        QtéUnité:         0,
        QtéParCarton:     unitsPerCarton,
        PoidsCarton:      poidsParCarton,
        VolCarton:        volParCarton,
        FullCartons:      0,
        CartonsÀExpédier: 0,
        VolumeTotal:      0,
        PoidsTotal:       0,
        Refrigerer:       prod["Refrigerer"]
      });
      return;
    }

    // Nombre de cartons entiers et reste
    const fullCartons    = Math.floor(qtUnits / unitsPerCarton);
    const remainderUnits = qtUnits % unitsPerCarton;
    const cartonsToShip  = fullCartons + (remainderUnits > 0 ? 1 : 0);

    // Total poids/volume (par cartons)
    const totalWeight = cartonsToShip * poidsParCarton;
    const totalVolume = cartonsToShip * volParCarton;

    // Si reste > 0 ⇒ message carton incomplet
    if (remainderUnits > 0) {
      const missing = unitsPerCarton - remainderUnits;
      missingMessages.push(
        `Il manque ${missing} unité(s) de ${prod["Référence"]} pour remplir le dernier carton.`
      );
    }

    // Ajouter aux totaux réfrigéré ou sec
    if (isRefrig) {
      totalRefPds += totalWeight;
      totalRefVol += totalVolume;
    } else {
      totalDryPds += totalWeight;
      totalDryVol += totalVolume;
    }

    // Stocker dans rowsProduits pour l’export Excel
    lastResult.rowsProduits.push({
      Référence:        prod["Référence"],
      Nom:              prod["Product"] + " "+ prod["Presentation"] + " " + prod["Unnamed: 3"],
      QtéUnité:         qtUnits,
      QtéParCarton:     unitsPerCarton,
      PoidsCarton:      parseFloat(poidsParCarton.toFixed(3)),
      VolCarton:        parseFloat(volParCarton.toFixed(6)),
      FullCartons:      fullCartons,
      CartonsÀExpédier: cartonsToShip,
      VolumeTotal:      parseFloat(totalVolume.toFixed(6)),
      PoidsTotal:       parseFloat(totalWeight.toFixed(3)),
      Refrigerer:       prod["Refrigerer"]
    });
  });

  const totalVolAll = totalRefVol + totalDryVol;
  const totalPdsAll = totalRefPds + totalDryPds;

  // Si aucune quantité totale (ni réfrigéré, ni sec)
  if (totalVolAll === 0 && totalPdsAll === 0) {
    afficherMessage({ html: `<div class="message"><em>Aucune quantité saisie.</em></div>` });
    lastResult.ref = null;
    lastResult.dry = null;
    return;
  }

  // Initialisation
  lastResult.ref = null;
  lastResult.dry = null;
  let htmlResultat = "";
  let resteVolRef = 0, restePdsRef = 0;

  // 3) Partie RÉFRIGÉRÉE
  if (totalRefVol > 0 || totalRefPds > 0) {
    // Filtrer les conteneurs réfrigérés (ex : “TC20R”, “TC40R”, “TC40CHR”)
    const contRef = conteneurs.filter(c => {
      const code = (c["NAME "] || "").trim();
      return code === "TC20R" || code === "TC40R" || code === "TC40CHR";
    });
    const resRef = findOptimalContainers(totalRefVol, totalRefPds, contRef);

    lastResult.ref = {
      totalVol: totalRefVol,
      totalPds: totalRefPds,
      resultat: resRef
    };

    resteVolRef = resRef.resteVolume;
    restePdsRef = resRef.restePoids;

    htmlResultat += formatResultMessage(
      "Conteneur(s) réfrigéré(s) pour produits réfrigérés",
      totalRefVol,
      totalRefPds,
      resRef
    );
  }

  // 4) Placer le sec dans l’espace restant Réfrigéré
  let remainDryVol = totalDryVol;
  let remainDryPds = totalDryPds;
  if ((totalRefVol > 0 || totalRefPds > 0) && (totalDryVol > 0 || totalDryPds > 0)) {
    if (remainDryVol <= resteVolRef && remainDryPds <= restePdsRef) {
      htmlResultat += `
        <div class="message categorie">
          <div class="message-item">Tous les cartons de produits non réfrigérés tiennent dans l’espace restant des conteneurs réfrigérés.</div>
        </div>
      `;
      remainDryVol = 0;
      remainDryPds = 0;
    } else {
      remainDryVol -= resteVolRef;
      remainDryPds -= restePdsRef;
      remainDryVol = Math.max(0, remainDryVol);
      remainDryPds = Math.max(0, remainDryPds);
    }
  }

  // 5) Si du sec reste, conteneurs secs
  if (remainDryVol > 0 || remainDryPds > 0) {
    const contDry = conteneurs.filter(c => {
      const code = (c["NAME "] || "").trim();
      return code !== "TC20R" && code !== "TC40R" && code !== "TC40CHR";
    });
    const resDry = findOptimalContainers(remainDryVol, remainDryPds, contDry);
    lastResult.dry = {
      totalVol: remainDryVol,
      totalPds: remainDryPds,
      resultat: resDry
    };
    htmlResultat += formatResultMessage(
      "Conteneur(s) sec(s) pour produits non réfrigérés restants",
      remainDryVol,
      remainDryPds,
      resDry
    );
  } else {
    if (lastResult.ref) {
      lastResult.dry = null;
      htmlResultat += `
        <div class="message categorie">
          <div class="message-item">Aucun container sec requis (tout tient dans le(s) container(s) réfrigéré(s)).</div>
        </div>
      `;
    }
  }

  // 6) Ajouter les messages sur cartons incomplets
  if (missingMessages.length > 0) {
    htmlResultat += `
      <div class="message categorie">
        <div class="message-item titre">Cartons incomplets :</div>
    `;
    missingMessages.forEach(msg => {
      htmlResultat += `<div class="message-item">⚠️ ${msg}</div>`;
    });
    htmlResultat += `</div>`;
  }

  // 7) Affichage final
  afficherMessage({ html: htmlResultat });
}

/**
 * findOptimalContainers(totalVol, totalPds, availableContainers):
 * – totalVol, totalPds : besoins à couvrir
 * – availableContainers : array ({ "NAME ", "Capacite_plus_de_quatre", "Poids_max" })
 *
 * Renvoie { containers:[codes], capVolume, capPoids, resteVolume, restePoids } ou .error
 */
function findOptimalContainers(totalVol, totalPds, availableContainers) {
  // 1. Construire et trier la liste
  const list = availableContainers
    .map(c => ({
      code:   (c["NAME "] || "").trim(),
      volCap: parseFloat(c["Capacite_plus_de_quatre"]),
      pdsCap: parseFloat(c["Poids_max"])
    }))
    .filter(c => c.code && !isNaN(c.volCap) && !isNaN(c.pdsCap))
    .sort((a, b) => {
      if (a.volCap !== b.volCap) return a.volCap - b.volCap;
      return a.pdsCap - b.pdsCap;
    });

  // 2. Chercher un conteneur unique adapté
  let meilleurMono = null;
  for (let c of list) {
    if (c.volCap >= totalVol && c.pdsCap >= totalPds) {
      const wasteVol = c.volCap - totalVol;
      const wastePds = c.pdsCap - totalPds;
      if (
        !meilleurMono ||
        wasteVol < meilleurMono.wasteVol ||
        (wasteVol === meilleurMono.wasteVol && wastePds < meilleurMono.wastePds)
      ) {
        meilleurMono = { container: c, wasteVol, wastePds };
      }
    }
  }
  if (meilleurMono) {
    const c = meilleurMono.container;
    return {
      containers:   [c.code],
      capVolume:    c.volCap,
      capPoids:     c.pdsCap,
      resteVolume:  parseFloat((c.volCap - totalVol).toFixed(6)),
      restePoids:   parseFloat((c.pdsCap - totalPds).toFixed(3))
    };
  }

  // 3. Rechercher la meilleure paire (i ≤ j)
  let meilleurPair = null;
  for (let i = 0; i < list.length; i++) {
    for (let j = i; j < list.length; j++) {
      const c1 = list[i];
      const c2 = list[j];
      const volSum = c1.volCap + c2.volCap;
      const pdsSum = c1.pdsCap + c2.pdsCap;
      if (volSum >= totalVol && pdsSum >= totalPds) {
        const wasteVol = volSum - totalVol;
        const wastePds = pdsSum - totalPds;
        if (
          !meilleurPair ||
          wasteVol < meilleurPair.wasteVol ||
          (wasteVol === meilleurPair.wasteVol && wastePds < meilleurPair.wastePds)
        ) {
          meilleurPair = { pair: [c1, c2], wasteVol, wastePds };
        }
      }
    }
  }
  if (meilleurPair) {
    const [c1, c2] = meilleurPair.pair;
    return {
      containers:   [c1.code, c2.code],
      capVolume:    c1.volCap + c2.volCap,
      capPoids:     c1.pdsCap + c2.pdsCap,
      resteVolume:  parseFloat(((c1.volCap + c2.volCap) - totalVol).toFixed(6)),
      restePoids:   parseFloat(((c1.pdsCap + c2.pdsCap) - totalPds).toFixed(3))
    };
  }

  // 4. Sinon, multiplier le + grand
  if (list.length === 0) {
    return {
      containers:   [],
      capVolume:    0,
      capPoids:     0,
      resteVolume:  0,
      restePoids:   0,
      error:        "Aucun conteneur disponible dans cette catégorie."
    };
  }
  const largest = list[list.length - 1];
  const nbByVol = Math.ceil(totalVol  / largest.volCap);
  const nbByPds = Math.ceil(totalPds  / largest.pdsCap);
  const nbNeeded = Math.max(nbByVol, nbByPds);
  const totalCapVol = largest.volCap * nbNeeded;
  const totalCapPds = largest.pdsCap * nbNeeded;

  return {
    containers:   Array(nbNeeded).fill(largest.code),
    capVolume:    totalCapVol,
    capPoids:     totalCapPds,
    resteVolume:  parseFloat((totalCapVol - totalVol).toFixed(6)),
    restePoids:   parseFloat((totalCapPds - totalPds).toFixed(3))
  };
}

/**
 * formatResultMessage(titreCat, totalVol, totalPds, resultat) :
 * Génère le HTML d’une “catégorie” (R ou Sec) :
 * – titreCat : ex. “Conteneur(s) réfrigéré(s) …”
 * – totalVol/totalPds : besoins réels
 * – resultat : {containers, capVolume, capPoids, resteVolume, restePoids}
 */
function formatResultMessage(titreCat, totalVol, totalPds, resultat) {
  let html = `<div class="message categorie">`;
  html += `<div class="message-item titre">${titreCat} :</div>`;

  if (resultat.error) {
    html += `<div class="message-item">⚠️ ${resultat.error}</div>`;
  } else {
    const codes = resultat.containers.join(" + ");
    html += `<div class="message-item">📦 Conteneur(s) sélectionné(s) : <strong>${codes}</strong></div>`;
    html += `<div class="message-item">🔍 Capacité totale : <strong>${resultat.capVolume.toLocaleString("fr-FR", { minimumFractionDigits: 6 })} m³</strong> et <strong>${resultat.capPoids.toLocaleString("fr-FR", { minimumFractionDigits: 3 })} kg</strong></div>`;
    html += `<div class="message-item">⚖️ Besoins totaux : <strong>${totalVol.toLocaleString("fr-FR", { minimumFractionDigits: 6 })} m³</strong> et <strong>${totalPds.toLocaleString("fr-FR", { minimumFractionDigits: 3 })} kg</strong></div>`;
    html += `<div class="message-item">✅ Espace restant : <strong>${resultat.resteVolume.toLocaleString("fr-FR", { minimumFractionDigits: 6 })} m³</strong> et <strong>${resultat.restePoids.toLocaleString("fr-FR", { minimumFractionDigits: 3 })} kg</strong></div>`;
  }
  html += `</div>`;
  return html;
}

/**
 * afficherMessage({ html }) :
 * Injecte le HTML passé dans la div #message-resultat.
 */
function afficherMessage({ html }) {
  const zone = document.getElementById("message-resultat");
  zone.innerHTML = html;
}

/**
 * resetForm() :
 * – Remet toutes les quantités à zéro
 * – Réinitialise les inputs “Quantité par carton”, “Poids Carton”, “Volume Carton” et “Réfrigéré ?” à leurs valeurs JSON par défaut
 * – Vide la zone résultat
 */
function resetForm() {
  // Réinitialiser toutes les quantités à zéro et les cases réfrigéré + champs carton
  produits.forEach((prod, i) => {
    const inputQt             = document.getElementById(`quantite-${i}`);
    const inputQtCarton       = document.getElementById(`prod-qtparcarton-${i}`);
    const inputPoidsCarton    = document.getElementById(`prod-poidscarton-${i}`);
    const inputVolCarton      = document.getElementById(`prod-volcarton-${i}`);
    const checkboxRef         = document.getElementById(`prod-refrig-${i}`);

    if (inputQt)           inputQt.value = 0;
    if (inputQtCarton)     inputQtCarton.value = prod["Quantité par carton"];
    if (inputPoidsCarton)  inputPoidsCarton.value = parseFloat(prod["Poids Brut Carton"]).toLocaleString("en", { useGrouping: false, minimumFractionDigits: 3 });
    if (inputVolCarton)    inputVolCarton.value = parseFloat(prod["M3 par carton"]).toLocaleString("en", { useGrouping: false, minimumFractionDigits: 6 });
    if (checkboxRef)       checkboxRef.checked = prod["Refrigerer"] == 1;
  });

  // Vider la zone résultat
  document.getElementById("message-resultat").innerHTML = "";

  // Reset lastResult
  lastResult.ref = null;
  lastResult.dry = null;
  lastResult.rowsProduits = [];

  // Réafficher les tableaux pour remettre tout à jour
  genererTableProduits();
  genererTableConteneursOverrides();
}

/**
 * packProducts() :
 * Répartit unitairement les produits dans des conteneurs physiques
 * après qu’on ait sélectionné les conteneurs via findOptimalContainers.
 * On travaille ici par cartons : on soustrait du vol/poids du conteneur
 * la valeur “VolCarton” et “PoidsCarton” pour chaque carton placé.
 *
 * Retourne un tableau de conteneurs physiques avec leur contenu.
 */
function packProducts() {
  const containersPhysiques = [];
  const compteurCode = {};

  // Ajoute “count” exemplaires du conteneur codePad
  // en leur attribuant un suffixe unique (“-1”, “-2”, …).
  function addContainers(codePad, count) {
    const info = conteneurs.find(c => (c["NAME "] || "").trim() === codePad);
    if (!info) return;
    const volCap = parseFloat(info["Capacite_plus_de_quatre"]);
    const pdsCap = parseFloat(info["Poids_max"]);

    for (let i = 0; i < count; i++) {
      compteurCode[codePad] = (compteurCode[codePad] || 0) + 1;
      const index = compteurCode[codePad];
      const codeUnique = `${codePad}-${index}`;
      containersPhysiques.push({
        codeUnique: codeUnique,
        baseCode: codePad,
        remainingVol: volCap,
        remainingPds: pdsCap,
        items: {}
      });
    }
  }

  // 1) Ajouter conteneurs réfrigérés
  if (lastResult.ref && lastResult.ref.resultat) {
    lastResult.ref.resultat.containers.forEach(codePad =>
      addContainers(codePad, 1)
    );
  }
  // 2) Ajouter conteneurs secs
  if (lastResult.dry && lastResult.dry.resultat) {
    lastResult.dry.resultat.containers.forEach(codePad =>
      addContainers(codePad, 1)
    );
  }

  // 3) Construire la liste des produits à placer (avec QtéUnité et infos carton à jour)
  const produitsList = lastResult.rowsProduits
    .map(r => ({
      reference:      r.Référence,
      refrigerer:     r.Refrigerer,
      qtyUnits:       r.QtéUnité,
      unitsPerCarton: parseInt(r.QtéParCarton, 10),
      volCarton:      parseFloat(r.VolCarton),
      pdsCarton:      parseFloat(r.PoidsCarton)
    }))
    .filter(r => r.qtyUnits > 0);

  // 4) Placer les cartons réfrigérés dans conteneurs R
  produitsList
    .filter(p => p.refrigerer == 1)
    .forEach(p => {
      let unitsToPlace = p.qtyUnits;
      while (unitsToPlace > 0) {
        // On cherche un conteneur R avec assez de place pour un carton entier
        const idx = containersPhysiques.findIndex(c =>
          c.baseCode.endsWith("R") &&
          c.remainingVol >= p.volCarton &&
          c.remainingPds >= p.pdsCarton
        );
        if (idx < 0) {
          console.warn(`Impossible de placer ${p.reference} réfrigéré`);
          break;
        }
        const container = containersPhysiques[idx];
        container.remainingVol -= p.volCarton;
        container.remainingPds -= p.pdsCarton;
        // On stocke “unitsPerCarton” dans items
        container.items[p.reference] = (container.items[p.reference] || 0) + p.unitsPerCarton;
        unitsToPlace -= p.unitsPerCarton;
      }
    });

  // 5) Placer les cartons secs (d’abord dans R s’il reste de la place, sinon dans Sec)
  produitsList
    .filter(p => p.refrigerer == 0)
    .forEach(p => {
      let unitsToPlace = p.qtyUnits;
      while (unitsToPlace > 0) {
        // Essayer un conteneur R
        let idx = containersPhysiques.findIndex(c =>
          c.baseCode.endsWith("R") &&
          c.remainingVol >= p.volCarton &&
          c.remainingPds >= p.pdsCarton
        );
        if (idx < 0) {
          // Sinon, chercher un Sec
          idx = containersPhysiques.findIndex(c =>
            !c.baseCode.endsWith("R") &&
            c.remainingVol >= p.volCarton &&
            c.remainingPds >= p.pdsCarton
          );
        }
        if (idx < 0) {
          console.warn(`Impossible de placer ${p.reference} non-réfrigéré`);
          break;
        }
        const container = containersPhysiques[idx];
        container.remainingVol -= p.volCarton;
        container.remainingPds -= p.pdsCarton;
        container.items[p.reference] = (container.items[p.reference] || 0) + p.unitsPerCarton;
        unitsToPlace -= p.unitsPerCarton;
      }
    });

  return containersPhysiques;
}

/**
 * downloadExcel() :
 * – Vérifie qu’un calcul existe (lastResult.ref ou lastResult.dry)
 * – Construit trois feuilles dans un même classeur :
 *     1) “Containers”  : résumé global par catégorie (R / Sec)
 *     2) “Produits”    : détail produit (QtéUnité, QtéParCarton, n° cartons, totaux)
 *     3) “Breakdown”   : détail par conteneur physique
 * – Télécharge le fichier “composition_containers.xlsx”.
 */
function downloadExcel() {
  if (!lastResult.ref && !lastResult.dry) {
    alert("Aucun résultat à exporter. Veuillez d'abord faire un calcul.");
    return;
  }

  // 1) Feuille “Containers”
  const dataContainers = [];
  if (lastResult.ref) {
    const { totalVol, totalPds, resultat } = lastResult.ref;
    const usedVol = parseFloat((resultat.capVolume - resultat.resteVolume).toFixed(6));
    const usedPds = parseFloat((resultat.capPoids - resultat.restePoids).toFixed(3));
    dataContainers.push({
      Category:        "Réfrigéré",
      Containers:      resultat.containers.join(" + "),
      CapVolume:       resultat.capVolume,
      CapPoids:        resultat.capPoids,
      UsedVolume:      usedVol,
      UsedPoids:       usedPds,
      RemainingVolume: resultat.resteVolume,
      RemainingPoids:  resultat.restePoids
    });
  }
  if (lastResult.dry) {
    const { totalVol, totalPds, resultat } = lastResult.dry;
    const usedVol = parseFloat((resultat.capVolume - resultat.resteVolume).toFixed(6));
    const usedPds = parseFloat((resultat.capPoids - resultat.restePoids).toFixed(3));
    dataContainers.push({
      Category:        "Sec",
      Containers:      resultat.containers.join(" + "),
      CapVolume:       resultat.capVolume,
      CapPoids:        resultat.capPoids,
      UsedVolume:      usedVol,
      UsedPoids:       usedPds,
      RemainingVolume: resultat.resteVolume,
      RemainingPoids:  resultat.restePoids
    });
  }

  // 2) Feuille “Produits”
  const dataProduits = lastResult.rowsProduits.map(row => ({
    Référence:        row.Référence,
    Nom:              row.Nom,
    QtéUnité:         row.QtéUnité,
    QtéParCarton:     row.QtéParCarton,
    PoidsCarton:      row.PoidsCarton,
    VolCarton:        row.VolCarton,
    FullCartons:      row.FullCartons,
    CartonsÀExpédier: row.CartonsÀExpédier,
    VolumeTotal:      row.VolumeTotal,
    PoidsTotal:       row.PoidsTotal,
    Réfrigéré:        row.Refrigerer
  }));

  // 3) Feuille “Breakdown” (via packProducts)
  const containersPhysiques = packProducts();
  const dataBreakdown = [];
  containersPhysiques.forEach(container => {
    const codeCont = container.codeUnique;
    for (const [ref, totalUnitsInContainer] of Object.entries(container.items)) {
      dataBreakdown.push({
        ContainerCode: codeCont,
        Référence:     ref,
        Quantité:      totalUnitsInContainer
      });
    }
    if (Object.keys(container.items).length === 0) {
      dataBreakdown.push({
        ContainerCode: codeCont,
        Référence:     "",
        Quantité:      0
      });
    }
  });

  // 4) Création du classeur et des feuilles
  const wb = XLSX.utils.book_new();

  const wsContainers = XLSX.utils.json_to_sheet(dataContainers, {
    header: [
      "Category",
      "Containers",
      "CapVolume",
      "CapPoids",
      "UsedVolume",
      "UsedPoids",
      "RemainingVolume",
      "RemainingPoids"
    ]
  });
  XLSX.utils.book_append_sheet(wb, wsContainers, "Containers");

  const wsProduits = XLSX.utils.json_to_sheet(dataProduits, {
    header: [
      "Référence",
      "Nom",
      "QtéUnité",
      "QtéParCarton",
      "PoidsCarton",
      "VolCarton",
      "FullCartons",
      "CartonsÀExpédier",
      "VolumeTotal",
      "PoidsTotal",
      "Réfrigéré"
    ]
  });
  XLSX.utils.book_append_sheet(wb, wsProduits, "Produits");

  const wsBreakdown = XLSX.utils.json_to_sheet(dataBreakdown, {
    header: [
      "ContainerCode",
      "Référence",
      "Quantité"
    ]
  });
  XLSX.utils.book_append_sheet(wb, wsBreakdown, "Breakdown");

  XLSX.writeFile(wb, "composition_containers.xlsx");
}
