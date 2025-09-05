
import React, { useState, useEffect, useCallback } from 'react';
import { AddRecordForm } from './components/AddRecordForm';
import { DuplicateHealth } from './components/DuplicateHealth';
import { Message } from './components/Message';
import { addRecordAndCheckDuplicates, getDuplicateHealthData } from './services/googleScript';
import type { AddRecordFormData, AddRecordResponse, DuplicateHealthData } from './types';
import { MessageType } from './types';

const App: React.FC = () => {
  const [healthData, setHealthData] = useState<DuplicateHealthData | null>(null);
  const [isLoadingHealth, setIsLoadingHealth] = useState<boolean>(true);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [message, setMessage] = useState<{ text: string; subtext?: string; type: MessageType }>({ text: '', type: MessageType.None });

  const fetchHealthData = useCallback(async () => {
    setIsLoadingHealth(true);
    try {
      const data = await getDuplicateHealthData();
      setHealthData(data);
    } catch (error) {
      setMessage({ text: 'Failed to load duplicate health data.', type: MessageType.Error });
      console.error(error);
    } finally {
      setIsLoadingHealth(false);
    }
  }, []);

  useEffect(() => {
    fetchHealthData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Auto-clear message after a delay
  useEffect(() => {
    if (message.text) {
      const timer = setTimeout(() => {
        setMessage({ text: '', type: MessageType.None });
      }, 7000); // 7 seconds
      
      // Cleanup timer if a new message is set before the old one clears
      return () => clearTimeout(timer);
    }
  }, [message]);

  const handleAddRecord = async (formData: AddRecordFormData): Promise<void> => {
    setIsSubmitting(true);
    setMessage({ text: '', type: MessageType.None });
    try {
      const response: AddRecordResponse = await addRecordAndCheckDuplicates(formData);
      
      if (response.success) {
        setMessage({ text: response.message, type: MessageType.Success });
        fetchHealthData(); // On success, fetch fresh (likely good) data
      } else {
        if (response.duplicatesFound && response.duplicatesFound > 0) {
          setMessage({
            text: response.message,
            subtext: 'USE STAR LABS',
            type: MessageType.Warning,
          });
          // As requested, update the health stats to show the "red" 30% state
          setHealthData(prevData => {
              const totalChecked = (prevData?.totalChecked ?? 0) + 1;
              // To make the percentage exactly 30, calculate the required duplicates
              const totalDuplicates = Math.ceil(totalChecked * 0.30);
              return {
                  percentage: 30,
                  totalChecked: totalChecked,
                  totalDuplicates: totalDuplicates,
              };
          });
        } else {
          setMessage({ text: response.message, type: MessageType.Error });
        }
      }
    } catch (error: any) {
      setMessage({ text: error.message || 'An unexpected error occurred.', type: MessageType.Error });
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-800 text-gray-100 p-4 font-sans">
      <div className="max-w-md mx-auto bg-gray-900 shadow-2xl rounded-lg overflow-hidden">
        <header className="p-5 bg-gray-700 border-b border-gray-600">
          <h1 className="text-2xl font-bold text-center text-cyan-400">BHG DeDuper</h1>
        </header>
        <main className="p-6 space-y-6">
          <DuplicateHealth data={healthData} isLoading={isLoadingHealth} />
          <div className="h-12 flex items-center justify-center">
             {message.text && <Message text={message.text} subtext={message.subtext} type={message.type} />}
          </div>
          <AddRecordForm onSubmit={handleAddRecord} isSubmitting={isSubmitting} />
        </main>
      </div>
    </div>
  );
};

export default App;