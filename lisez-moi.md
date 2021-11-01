# Windows - Script utilities

![Banner](./banner.svg)

Si, comme moi, vous devez travailler sous Windows et qu'il vous faut préparer une migration MS Access -> SQL Server, les utilitaires de ce dépôt pourront vous aider.

Actuellement (Nov 2017), les scripts proposés permettent de :

- Faire l'inventaire des fichiers existant sur un disque (local ou réseaux) (le scan se fait sur une ou plusieurs extensions; donc pas forcément uniquement MS Access)
- Faire l'inventaire des bases de données Access : liste des tables qui sont dans les DBs avec quelques propriétés comme le nom de la table, le fait qu'il s'agit ou pas d'une table locale et si c'est une table externe, nom de la DB d'origine ainsi que nom de la table
- Exportation automatisée du code VBA des bases de données sous la forme de fichiers textes (=> backup et versionning)
- Scan de la liste des champs de la base de données avec, pour les champs texte et memo, la taille max du champs ainsi que la taille de la plus petite et de la plus grande information stockée (la taille du champs pourrait par exemple être 255 alors que la plus petite info stockée est de 3 lettres et la plus longue de 10 lettres => il n'est vraiment pas utile de laisser le champs sur 255)
- Suppression d'un préfixe : si le nom des tables commencent par p.ex. "dbo\_", le script peut supprimer le préfixe.

La classe MSAccess permet donc de générer, de manière automatisée, un inventaire des bases de données, des tables, des champs, des liens entre bases de données et quelques fonctions de maintenance; cela en vue d'une migration vers SQL Serveur; p.ex.
