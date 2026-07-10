'use client';

import { useState } from 'react';

type ConfiguratorUIProps = {
  colors: Record<string, string>;
  setColor: (part: string, color: string) => void;
  basePrice: number;
};

const OPTIONS = {
  seat: {
    title: 'Seat Material',
    options: [
      { name: 'Light Grey', hex: '#e0e0e0', priceAddon: 0 },
      { name: 'Charcoal', hex: '#333333', priceAddon: 0 },
      { name: 'Navy Blue', hex: '#1a365d', priceAddon: 20 },
      { name: 'Forest Green', hex: '#276749', priceAddon: 20 },
      { name: 'Premium Leather', hex: '#742a2a', priceAddon: 80 },
    ]
  },
  backrest: {
    title: 'Backrest Finish',
    options: [
      { name: 'Charcoal', hex: '#333333', priceAddon: 0 },
      { name: 'Light Oak', hex: '#d4a373', priceAddon: 10 },
      { name: 'Dark Walnut', hex: '#5c4033', priceAddon: 15 },
      { name: 'White Ash', hex: '#ffffff', priceAddon: 10 },
    ]
  },
  legs: {
    title: 'Leg Style',
    options: [
      { name: 'Matte Black', hex: '#111111', priceAddon: 0 },
      { name: 'Chrome', hex: '#e8e8e8', priceAddon: 25 },
      { name: 'Solid Brass', hex: '#b5a642', priceAddon: 45 },
    ]
  },
};

const PARTS = Object.keys(OPTIONS) as Array<keyof typeof OPTIONS>;

export default function ConfiguratorUI({ colors, setColor, basePrice }: ConfiguratorUIProps) {
  const [currentStep, setCurrentStep] = useState(0);

  const activePart = PARTS[currentStep];
  const partData = OPTIONS[activePart];

  // Calculate total price based on selected colors
  const totalPrice = basePrice + PARTS.reduce((total, part) => {
    const selectedHex = colors[part];
    const option = OPTIONS[part].options.find(o => o.hex === selectedHex);
    return total + (option?.priceAddon || 0);
  }, 0);

  return (
    <>
      <div className="top-action-bar">
        <div className="brand-logo">Cognaq Studio</div>
        <div className="action-buttons">
          <button className="btn-icon">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line></svg>
            View in AR
          </button>
          <button className="btn-icon">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="18" cy="5" r="3"></circle><circle cx="6" cy="12" r="3"></circle><circle cx="18" cy="19" r="3"></circle><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"></line><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"></line></svg>
            Share
          </button>
        </div>
      </div>

      <div className="floating-panel">
        <div className="panel-header">
          <h2>Modular Chair</h2>
          <p>Customize your product</p>
        </div>

        <div className="steps">
          {PARTS.map((_, index) => (
            <div 
              key={index} 
              className={`step-indicator ${index <= currentStep ? 'active' : ''}`} 
            />
          ))}
        </div>

        <div className="step-content">
          <div className="step-title">
            <span>Step {currentStep + 1} of {PARTS.length}</span>
            <span className="part-name">{partData.title}</span>
          </div>

          <div className="color-grid">
            {partData.options.map((option) => (
              <div key={option.hex} className="color-swatch-container">
                <button
                  className={`color-swatch ${colors[activePart] === option.hex ? 'active' : ''}`}
                  style={{ backgroundColor: option.hex }}
                  onClick={() => setColor(activePart, option.hex)}
                  title={`${option.name} ${option.priceAddon > 0 ? `(+$${option.priceAddon})` : ''}`}
                  aria-label={option.name}
                />
                <span className="swatch-label">{option.name}</span>
              </div>
            ))}
          </div>

          <div className="step-nav">
            <button 
              className="nav-btn" 
              disabled={currentStep === 0} 
              onClick={() => setCurrentStep(prev => prev - 1)}
            >
              ← Previous
            </button>
            <button 
              className="nav-btn" 
              disabled={currentStep === PARTS.length - 1} 
              onClick={() => setCurrentStep(prev => prev + 1)}
            >
              Next →
            </button>
          </div>
        </div>
        
        <div className="panel-footer">
          <div className="price-row">
            <span className="price-label">Total Price</span>
            <span className="price-value">${totalPrice}</span>
          </div>
          <button className="checkout-btn">Add to Cart</button>
        </div>
      </div>
    </>
  );
}
