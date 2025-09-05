
import React from 'react';
import { MessageType } from '../types';

interface MessageProps {
  text: string;
  subtext?: string;
  type: MessageType;
}

export const Message: React.FC<MessageProps> = ({ text, type, subtext }) => {
  if (!text) return null;

  const baseClasses = 'text-sm font-medium text-center p-2 rounded-md w-full';
  let typeClasses = '';

  switch (type) {
    case MessageType.Success:
      typeClasses = 'bg-green-900/50 text-green-300';
      break;
    case MessageType.Error:
      typeClasses = 'bg-red-900/50 text-red-300';
      break;
    case MessageType.Warning:
      typeClasses = 'bg-orange-800/70 text-orange-300';
      break;
    default:
      return null;
  }

  return (
    <div className={`${baseClasses} ${typeClasses}`} role="alert">
      <p>{text}</p>
      {subtext && <p className="text-red-400 font-bold mt-1">{subtext}</p>}
    </div>
  );
};
