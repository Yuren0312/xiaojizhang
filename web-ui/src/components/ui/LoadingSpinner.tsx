import React from 'react';

interface LoadingSpinnerProps {
  size?: 'sm' | 'md' | 'lg';
  text?: string;
  className?: string;
}

const sizeClasses = {
  sm: 'w-4 h-4 border-2',
  md: 'w-8 h-8 border-[3px]',
  lg: 'w-12 h-12 border-4',
};

const LoadingSpinner: React.FC<LoadingSpinnerProps> = ({
  size = 'md',
  text,
  className = '',
}) => {
  return (
    <div className={`flex flex-col items-center justify-center gap-3 ${className}`}>
      <div
        className={`${sizeClasses[size]} rounded-full border-primary-200 border-t-primary-500 
                    animate-spin`}
        role="status"
        aria-label="加载中"
      />
      {text && (
        <p className="text-sm text-text-secondary animate-pulse-glow">{text}</p>
      )}
    </div>
  );
};

export default LoadingSpinner;
