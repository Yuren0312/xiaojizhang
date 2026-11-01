import React from 'react';

const Footer: React.FC = () => {
  return (
    <footer className="bg-primary-900 text-primary-100 text-xs text-center py-3 mt-auto">
      <div className="max-w-[1200px] mx-auto px-6 flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-4">
        <span>© 小虾记账 v1.0</span>
        <span className="hidden sm:inline">|</span>
        <span>小企业会计准则</span>
        <span className="hidden sm:inline">|</span>
        <span>账套数据安全存储</span>
      </div>
    </footer>
  );
};

export default Footer;
