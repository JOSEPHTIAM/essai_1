import { DashboardData } from '../models/dashboard.model';

export async function getDashboardData(): Promise<DashboardData> {
  // Simulation d'un délai réseau (comme un appel API ou Supabase)
  // Plus tard, vous remplacerez ceci par un appel direct à votre client Supabase.
  await new Promise((resolve) => setTimeout(resolve, 100));

  return {
    stats: [
      {
        id: '1',
        title: 'Utilisateurs Actifs',
        value: '2,420',
        change: '+12.5%',
        isPositive: true,
      },
      {
        id: '2',
        title: 'Revenus Mensuels',
        value: '14,850 €',
        change: '+8.2%',
        isPositive: true,
      },
      {
        id: '3',
        title: 'Taux de Conversion',
        value: '3.6%',
        change: '-0.4%',
        isPositive: false,
      },
      {
        id: '4',
        title: 'Projets en Cours',
        value: '12',
        change: '+4',
        isPositive: true,
      },
    ],
    activities: [
      {
        id: '1',
        user: 'Alice Martin',
        action: 'a créé un nouveau projet',
        target: 'Refonte API',
        time: 'Il y a 5 min',
        status: 'Succès',
      },
      {
        id: '2',
        user: 'Thomas Dubois',
        action: 'a mis à jour son profil',
        target: 'Paramètres',
        time: 'Il y a 25 min',
        status: 'Succès',
      },
      {
        id: '3',
        user: 'Sophie Bernard',
        action: 'a initié un paiement',
        target: 'Facture #1042',
        time: 'Il y a 1 heure',
        status: 'En attente',
      },
      {
        id: '4',
        user: 'Marc Leroy',
        action: 'a supprimé un accès',
        target: 'Serveur Staging',
        time: 'Il y a 3 heures',
        status: 'Succès',
      },
    ],
  };
}