import React, { useState } from 'react';

interface Props { variant?: 'icon' | 'full' | 'horizontal'; size?: 'sm' | 'md' | 'lg'; subtitleText?: string; className?: string; }
// box size of the logo picture (it is resized to fit inside, keeping its shape)
const px = { sm: 36, md: 48, lg: 72 } as const;

/** The logo picture comes from this web link. To change the logo, replace the link below. */
export const LOGO_URL =
  'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTWAHEFyU2OmJyl65dzSmywXNmu_aLIxVZ9Z6bJ5H5Ph2O6QoDAgdWD9Vtc&s=10';

// If the link stops working, a simple "F2" badge is shown instead.
export const F2FintechLogo: React.FC<Props> = ({ variant = 'icon', size = 'md', subtitleText, className = '' }) => {
  const s = px[size];
  const [failed, setFailed] = useState(false);
  const mark = failed ? (
    <span style={{ width: s, height: s }} className="inline-flex items-center justify-center rounded-xl bg-blue-600 text-white font-black">F2</span>
  ) : (
    <img
      src={LOGO_URL}
      referrerPolicy="no-referrer"
      alt="F2 Fintech"
      onError={() => setFailed(true)}
      style={{ width: s, height: s }}
      className="object-contain rounded-lg shrink-0"
    />
  );
  if (variant === 'icon') return <span className={`inline-flex ${className}`}>{mark}</span>;
  return (
    <span className={`inline-flex ${variant === 'full' ? 'flex-col items-center' : 'items-center'} gap-2 ${className}`}>
      {mark}
      <span className="leading-tight text-left">
        <span className="block font-black text-slate-900">F2 Fintech</span>
        {subtitleText && <span className="block text-[11px] text-slate-500">{subtitleText}</span>}
      </span>
    </span>
  );
};
