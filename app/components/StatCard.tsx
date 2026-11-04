interface StatCardProps {
  title: string;
  value: string;
  change: string;
  isPositive: boolean;
}

export default function StatCard({ title, value, change, isPositive }: StatCardProps) {
  return (
    <div className="bg-white dark:bg-black border border-zinc-200 dark:border-zinc-800 rounded-2xl p-6 shadow-sm hover:border-indigo-500 dark:hover:border-indigo-500 transition-all">
      <p className="text-sm font-medium text-zinc-500 dark:text-zinc-400">{title}</p>
      <div className="flex items-baseline justify-between mt-4">
        <h3 className="text-2xl font-bold text-black dark:text-white tracking-tight">{value}</h3>
        <span
          className={`text-xs font-semibold px-2.5 py-1 rounded-full ${
            isPositive
              ? 'bg-indigo-50 text-indigo-700 dark:bg-indigo-950/50 dark:text-indigo-400'
              : 'bg-zinc-100 text-zinc-600 dark:bg-zinc-800 dark:text-zinc-300'
          }`}
        >
          {change}
        </span>
      </div>
    </div>
  );
}