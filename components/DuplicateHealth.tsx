
import React from 'react';
import type { DuplicateHealthData } from '../types';

interface DuplicateHealthProps {
  data: DuplicateHealthData | null;
  isLoading: boolean;
}

const getHealthColor = (percentage: number): string => {
  if (percentage <= 15) return 'text-green-400';
  if (percentage < 30) return 'text-yellow-400';
  return 'text-red-400';
};

const getHealthBarColor = (percentage: number): string => {
  if (percentage <= 15) return 'bg-green-500';
  if (percentage < 30) return 'bg-yellow-500';
  return 'bg-red-500';
};

export const DuplicateHealth: React.FC<DuplicateHealthProps> = ({ data, isLoading }) => {
  if (isLoading) {
    return (
      <div className="bg-gray-800 p-4 rounded-lg animate-pulse">
        <div className="h-4 bg-gray-700 rounded w-3/4 mb-3"></div>
        <div className="h-8 bg-gray-700 rounded w-1/4 mb-4"></div>
        <div className="h-2 bg-gray-700 rounded"></div>
      </div>
    );
  }

  if (!data) {
    return (
        <div className="bg-gray-800 p-4 rounded-lg text-center text-gray-400">
            No health data available.
        </div>
    );
  }

  const { percentage } = data;
  const colorClass = getHealthColor(percentage);
  const barColorClass = getHealthBarColor(percentage);
  
  return (
    <div className="bg-gray-800 p-4 rounded-lg border border-gray-700">
      <h2 className="text-lg font-semibold text-gray-200 mb-2">Duplicate Health</h2>
      <p className="text-xs text-gray-400 mb-3">Percentage of duplicates since previous Monday.</p>
      <div className="flex items-center gap-4">
        <p className={`text-4xl font-bold ${colorClass}`}>{percentage}%</p>
        <div className="w-full">
            <div className="flex justify-between text-xs text-gray-400 mb-1">
                <span>{data.totalDuplicates} Duplicates</span>
                <span>{data.totalChecked} Total</span>
            </div>
            <div className="w-full bg-gray-700 rounded-full h-2.5">
                <div 
                    className={`${barColorClass} h-2.5 rounded-full transition-all duration-500 ease-out`} 
                    style={{ width: `${percentage}%` }}
                ></div>
            </div>
        </div>
      </div>
    </div>
  );
};
