'use client';

import { useState } from 'react';

type ConfiguratorUIProps = {
  colors: Record<string, string>;
  setColor: (part: string, color: string) => void;
};

const OPTIONS = {
  seat: [
    { name: 'Light Grey', hex: '#e0e0e0' },
    { name: 'Charcoal', hex: '#333333' },
    { name: 'Navy Blue', hex: '#1a365d' },
    { name: 'Forest Green', hex: '#276749' },
    { name: 'Burgundy', hex: '#742a2a' },
  ],
  backrest: [
    { name: 'Charcoal', hex: '#333333' },
    { name: 'Light Wood', hex: '#d4a373' },
    { name: 'Dark Wood', hex: '#5c4033' },
    { name: 'White', hex: '#ffffff' },
  ],
  legs: [
    { name: 'Matte Black', hex: '#111111' },
    { name: 'Chrome', hex: '#e8e8e8' },
    { name: 'Brass', hex: '#b5a642' },
  ],
};

export default function ConfiguratorUI({ colors, setColor }: ConfiguratorUIProps) {
  const [activeTab, setActiveTab] = useState<keyof typeof OPTIONS>('seat');

  return (
    <div className="configurator-sidebar">
      <div className="sidebar-header">
        <h1>Modular Chair</h1>
        <p>Customize your product</p>
      </div>

      <div className="tabs">
        {(Object.keys(OPTIONS) as Array<keyof typeof OPTIONS>).map((part) => (
          <button
            key={part}
            className={`tab-btn ${activeTab === part ? 'active' : ''}`}
            onClick={() => setActiveTab(part)}
          >
            {part.charAt(0).toUpperCase() + part.slice(1)}
          </button>
        ))}
      </div>

      <div className="options-container">
        <h3>Select Color</h3>
        <div className="color-grid">
          {OPTIONS[activeTab].map((option) => (
            <button
              key={option.hex}
              className={`color-swatch ${colors[activeTab] === option.hex ? 'active' : ''}`}
              style={{ backgroundColor: option.hex }}
              onClick={() => setColor(activeTab, option.hex)}
              title={option.name}
              aria-label={option.name}
            />
          ))}
        </div>
      </div>
      
      <div className="sidebar-footer">
        <div className="price">$299.00</div>
        <button className="add-to-cart-btn">Add to Cart</button>
      </div>
    </div>
  );
}
