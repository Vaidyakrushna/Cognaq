'use client';

import { useState } from 'react';
import Scene from '@/components/Scene';
import ConfiguratorUI from '@/components/ConfiguratorUI';

export default function Home() {
  const [colors, setColors] = useState<Record<string, string>>({
    seat: '#e0e0e0',
    backrest: '#333333',
    legs: '#111111'
  });

  const handleSetColor = (part: string, color: string) => {
    setColors(prev => ({
      ...prev,
      [part]: color
    }));
  };

  return (
    <main className="app-container">
      {/* 3D Viewer */}
      <Scene colors={colors} />
      
      {/* UI Overlay */}
      <ConfiguratorUI colors={colors} setColor={handleSetColor} />
    </main>
  );
}
