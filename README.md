# Kit Écriture manuscrite pour Word

Ce kit sert à créer dans **Microsoft Word** des documents qui ressemblent à une écriture faite à la main : fiches pédagogiques, cours, exercices, avec des **équations mathématiques** et des **formules chimiques**. Le texte reste entièrement modifiable.

Il fonctionne avec Word pour Windows (Microsoft 365, Word 2016, 2019, 2021 ou plus récent).

---

## ⬇️ Télécharger le kit

**[Télécharger Kit-Ecriture-Manuscrite-Word.zip](Kit-Ecriture-Manuscrite-Word.zip?raw=1)** (environ 1,8 Mo).

Décompressez-le (clic droit → **Extraire tout**), puis suivez les étapes ci-dessous ou ouvrez `LISEZMOI.pdf`.

---

## Contenu du kit

| Dossier | Contenu | À quoi ça sert |
|---|---|---|
| `1-Police` | `PlaypenManuscrit-Regular.ttf` | La police d'écriture manuscrite, à installer |
| `1-Police` | `OFL.txt` | La licence libre de la police |
| `2-Macro` | `EquationsManuscrites.bas` | Une macro qui met toutes les équations en écriture manuscrite en un clic |
| `3-Modele` | `Modele_manuscrit.docx` | Un document vierge déjà réglé : texte bleu, titres rouges et verts |
| `4-Exemples` | `Exemple_fiche_pedagogique.docx` | Une fiche pédagogique avec tableau et formules chimiques |
| `4-Exemples` | `Exemple_equations.docx` | Un cours de maths avec fractions, racines et exposants |

---

## Étape 1 : installer la police (une seule fois)

1. Ouvrez le dossier `1-Police`.
2. Double-cliquez sur **PlaypenManuscrit-Regular.ttf**.
3. Cliquez sur **Installer**.
4. Si Word était ouvert, fermez-le et rouvrez-le.

La police **Playpen Manuscrit** apparaît maintenant dans la liste des polices de Word.

> Les documents du kit contiennent déjà la police. Ils s'affichent donc correctement même sans l'installer. L'installation sert à **écrire de nouveaux textes** avec cette police.

---

## Étape 2 : installer la macro des équations (une seule fois)

1. Ouvrez Word, puis appuyez sur **Alt + F11** : l'éditeur de macros s'ouvre.
2. Dans la colonne de gauche, cliquez sur **Normal**. La macro sera ainsi disponible dans tous vos documents.
3. Menu **Fichier → Importer un fichier…**, puis choisissez `2-Macro/EquationsManuscrites.bas`.
4. Fermez l'éditeur de macros.

**Pour avoir un bouton (facultatif) :**
**Fichier → Options → Barre d'outils Accès rapide**. Dans « Choisir les commandes dans les catégories suivantes », choisissez **Macros**. Sélectionnez **EquationsManuscrites**, cliquez sur **Ajouter**, puis sur **OK**.

---

## Étape 3 : écrire un document

### Le plus simple : partir du modèle

Ouvrez `3-Modele/Modele_manuscrit.docx` et enregistrez-le sous un nouveau nom. Tout est déjà réglé :

- **Normal** : écriture manuscrite bleue ;
- **Titre** : titre de la leçon, en rouge, centré et souligné ;
- **Titre 1** : grands titres rouges en majuscules (PRÉREQUIS, MATÉRIEL…) ;
- **Titre 2** : sous-titres verts soulignés.

Les styles se choisissent dans l'onglet **Accueil**, zone **Styles**.

### Dans un document existant

Sélectionnez tout (**Ctrl + A**), puis choisissez la police **Playpen Manuscrit** et la couleur bleue.

---

## Étape 4 : les équations

1. Appuyez sur **Alt + =** pour insérer une équation.
2. Tapez-la, par exemple `x_1=(-b-\sqrt Δ)/(2a)`, puis appuyez sur **Espace** : Word la met en forme.
3. Quand toutes vos équations sont écrites, appuyez sur **Alt + F8**, choisissez **EquationsManuscrites** et cliquez sur **Exécuter** (ou cliquez sur votre bouton).

Toutes les équations passent en écriture manuscrite bleue.

**Bon à savoir :**

- Le signe racine (√) et la barre de fraction restent dans la police mathématique de Word. Word ne permet pas de les changer, mais l'ensemble reste cohérent.
- Après conversion, Word n'ajoute plus d'espaces automatiques autour de `=` et `+`. Tapez-les vous-même pour une écriture plus aérée.
- Vous pouvez relancer la macro autant de fois que vous voulez, par exemple après avoir ajouté de nouvelles équations.
- **Ne cherchez pas** la police dans *Options des équations → Police par défaut des fonctions mathématiques*. Cette liste n'affiche que les polices spéciales pour les maths (Cambria Math…). C'est normal, la macro contourne ce problème.

**Sans macro (à la main) :** sélectionnez l'équation, cliquez sur **« abc Texte normal »** dans l'onglet **Équation**, puis choisissez la police **Playpen Manuscrit** dans l'onglet **Accueil**.

---

## Étape 5 : les petites formules dans le texte

Pour CO₂, Al₂O₃, x², m³…, pas besoin de l'éditeur d'équations :

- **Indice** (le 2 de CO₂) : sélectionnez le chiffre, puis **Ctrl + =** ;
- **Exposant** (le 2 de x²) : sélectionnez le chiffre, puis **Ctrl + Maj + =**.

---

## Changer la couleur ou la police des équations

Ouvrez la macro (**Alt + F11**, puis Normal → Modules → EquationsManuscrites) et modifiez les deux lignes du début :

```vb
Private Const POLICE As String = "Playpen Manuscrit"
Private Const COULEUR As String = "1B3A8C"   ' bleu stylo
```

Quelques couleurs : noir `000000`, bleu stylo `1B3A8C`, rouge `D32F2F`, vert `2E7D32`.

---

## Partager un document

- **Avec quelqu'un qui a Word** : envoyez le `.docx`. La police est intégrée aux documents du kit. Pour vos propres documents, activez **Fichier → Options → Enregistrement → Incorporer les polices dans le fichier**.
- **Pour imprimer, ou pour quelqu'un sans Word** : **Fichier → Enregistrer sous → PDF**. C'est le plus sûr.

---

## Problèmes fréquents

| Problème | Solution |
|---|---|
| La police n'apparaît pas dans Word | Installez-la (étape 1), puis fermez et rouvrez Word. |
| Alt + F8 ne montre pas la macro | Refaites l'étape 2 en cliquant bien sur **Normal** avant d'importer. |
| « Les macros ont été désactivées » | Une macro importée dans Normal fonctionne normalement. Sinon : **Fichier → Options → Centre de gestion de la confidentialité → Paramètres des macros**. |
| Le texte s'affiche dans une autre police chez quelqu'un d'autre | Envoyez un PDF, ou activez l'incorporation des polices (voir « Partager un document »). |
| Je veux annuler la conversion des équations | **Ctrl + Z** juste après avoir lancé la macro. |

---

## Licence

- **Police** : *Playpen Manuscrit* est une version de **Playpen Sans** (© 2023 The Playpen Sans Project Authors, TypeTogether), réglée en graisse fine et renommée pour Word. Elle est distribuée sous la licence libre **SIL Open Font License 1.1** (voir `1-Police/OFL.txt`). Elle est gratuite, utilisable même pour des documents commerciaux, mais ne peut pas être vendue seule.
- **Macro, modèle et exemples** : libres d'utilisation et de partage.
