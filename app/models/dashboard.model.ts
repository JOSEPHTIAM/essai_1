export interface StatItem {
  id: string;
  title: string;
  value: string;
  change: string;
  isPositive: boolean;
}

export interface ActivityItem {
  id: string;
  user: string;
  action: string;
  target: string;
  time: string;
  status: 'Succès' | 'En attente' | 'Échec';
}

export interface DashboardData {
  stats: StatItem[];
  activities: ActivityItem[];
}