import React from 'react';
import Skeleton from './Skeleton';

const SkeletonCard: React.FC = () => {
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-6 space-y-4">
      <div className="flex items-center gap-3">
        <Skeleton variant="circular" width={48} height={48} />
        <div className="space-y-2 flex-1">
          <Skeleton variant="text" width="60%" height={20} />
          <Skeleton variant="text" width="40%" height={14} />
        </div>
      </div>
      <Skeleton variant="text" width="75%" height={14} />
      <Skeleton variant="rectangular" width="100%" height={40} />
    </div>
  );
};

export default SkeletonCard;
