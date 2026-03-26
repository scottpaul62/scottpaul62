# Macros SolidWorks

Trois macros VBA prêtes à l'emploi — aucun accès admin requis.

## Comment utiliser une macro

1. Télécharger le fichier `.bas` souhaité
2. Dans SolidWorks : **Outils > Macros > Exécuter...**
3. Sélectionner le fichier `.bas`
4. La macro se lance immédiatement

> Astuce : associer chaque macro à un bouton dans la barre d'outils via  
> **Outils > Personnaliser > Commandes > Macros**

---

## Macros disponibles

### 1. `SetupPiece.bas` — Propriétés standard

Remplit automatiquement les propriétés personnalisées :

| Propriété   | Saisie |
|-------------|--------|
| Référence   | Manuelle |
| Désignation | Manuelle |
| Matière     | Manuelle |
| Auteur      | Auto (login Windows) |
| Projet      | Manuelle |
| Révision    | Manuelle (défaut A) |
| Date        | Auto (aujourd'hui) |
| Statut      | Auto ("En cours") |

Ces propriétés alimentent directement le cartouche si votre template de plan est configuré avec les liens `$PRP:"Référence"` etc.

---

### 2. `CreerMiseEnPlan.bas` — Mise en plan automatique

Demande :
- Format (A4 / A3 / A2 / A1)
- Échelle (1:1, 1:2, 1:5, 2:1...)

Génère automatiquement :
- Vue **Face**
- Vue **Dessus**  
- Vue **Droite**
- Vue **Isométrique**

Sauvegarde automatiquement le `.SLDDRW` dans le même dossier que la pièce.

---

### 3. `NettoyerFichier.bas` — Nettoyage & diagnostic

- Détecte les **features en erreur** et les liste
- Identifie les **corps solides cachés** inutiles
- Force la **reconstruction** du modèle
- Supprime les **objets OLE** embarqués
- Sauvegarde et affiche un **rapport**

---

## Adapter les macros à votre entreprise

- **Noms des propriétés** : modifier les chaînes `"Référence"`, `"Désignation"`... pour correspondre à votre cartouche
- **Template de plan** : dans `CreerMiseEnPlan.bas`, modifier `sTemplate` pour pointer vers votre cartouche entreprise
- **Valeurs par défaut** : adapter les `InputBox` pour pré-remplir selon vos conventions
