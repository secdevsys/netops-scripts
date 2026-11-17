# Analyse de configuration Check Point

## Couches de politique

| # | Couche | Type | Règles | Accept | Drop | Inline | Désactivées | Sections |
|---|---|---|---|---|---|---|---|---|
| 1 | Network | ordonnée | 7 | 4 | 2 | 1 | 1 | 2 |
| 2 | Web_Control | inline | 3 | 1 | 2 | 0 | 0 | 0 |
| 3 | Application | ordonnée | 2 | 1 | 1 | 0 | 0 | 0 |

### Couche Network

Sections : Management, Internet

- règle 4 (Temp test) : Any → Any → Any en Accept
- règle 4 (Temp test) : aucune journalisation
- règle 4 (Temp test) : 0 hit (règle inutilisée ?)
- 1 règle(s) désactivée(s)

### Couche Web_Control

- règle 3.3 (Cleanup) : aucune journalisation

### Couche Application

- règle 2 (Default) : Any → Any → Any en Accept
- Pas de règle de nettoyage (Any/Any/Drop) en fin de couche (dernière : règle 2, Accept)

## Cluster CL-FW-PARIS

- Type : ClusterXL Gateway Cluster
- Adresse principale : 203.0.113.10
- Membres : fw-paris-01, fw-paris-02

### Interfaces virtuelles (VIP)

| Interface | VIP | Réseau | Topologie |
|---|---|---|---|
| eth1 | 203.0.113.10 | 203.0.113.0/28 | External (leads out to the Internet) |
| eth2 | 10.10.0.1 | 10.10.0.0/16 | Internal - Network defined by routes |
| eth3 | 172.16.50.1 | 172.16.50.0/24 | Internal - DMZ |
| eth4 | 192.168.100.1 | 192.168.100.0/24 | Internal - This network |

### Membre fw-paris-01

| Interface | IP | Réseau |
|---|---|---|
| eth1 | 203.0.113.11 | 203.0.113.0/28 |
| eth2 | 10.10.0.2 | 10.10.0.0/16 |
| eth3 | 172.16.50.2 | 172.16.50.0/24 |
| eth4 | 192.168.100.2 | 192.168.100.0/24 |
| Sync | 10.255.255.1 | 10.255.255.0/30 |

### Membre fw-paris-02

| Interface | IP | Réseau |
|---|---|---|
| eth1 | 203.0.113.12 | 203.0.113.0/28 |
| eth2 | 10.10.0.3 | 10.10.0.0/16 |
| eth3 | 172.16.50.3 | 172.16.50.0/24 |
| Sync | 10.255.255.2 | 10.255.255.0/30 |

### Réseaux

| Réseau | Rôle | VIP | Membres | Routeurs |
|---|---|---|---|---|
| 203.0.113.0/28 | Externe | 203.0.113.10 | fw-paris-01 203.0.113.11, fw-paris-02 203.0.113.12 | 203.0.113.1 → défaut |
| 10.10.0.0/16 | Interne | 10.10.0.1 | fw-paris-01 10.10.0.2, fw-paris-02 10.10.0.3 | 10.10.0.254 → 10.20.0.0/16, 10.30.0.0/16 |
| 172.16.50.0/24 | DMZ | 172.16.50.1 | fw-paris-01 172.16.50.2, fw-paris-02 172.16.50.3 |  |
| 192.168.100.0/24 | Interne | 192.168.100.1 | fw-paris-01 192.168.100.2 |  |
| 10.255.255.0/30 | Synchro |  | fw-paris-01 10.255.255.1, fw-paris-02 10.255.255.2 |  |

### Points d'attention

- Le membre fw-paris-02 n'a pas d'interface dans le réseau 192.168.100.0/24 (VIP 192.168.100.1)
