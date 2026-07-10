'use client';

import { Canvas } from '@react-three/fiber';
import { OrbitControls, Environment, ContactShadows } from '@react-three/drei';
import { Suspense } from 'react';

// A simple modular chair built with primitive shapes to demonstrate the configurator
function ChairModel({ colors }: { colors: Record<string, string> }) {
  return (
    <group position={[0, -0.5, 0]}>
      {/* Seat */}
      <mesh position={[0, 0.5, 0]} castShadow receiveShadow>
        <boxGeometry args={[1, 0.1, 1]} />
        <meshStandardMaterial color={colors.seat || '#e0e0e0'} roughness={0.7} />
      </mesh>
      
      {/* Backrest */}
      <mesh position={[0, 1.1, -0.45]} castShadow receiveShadow>
        <boxGeometry args={[1, 0.8, 0.1]} />
        <meshStandardMaterial color={colors.backrest || '#333333'} roughness={0.8} />
      </mesh>
      
      {/* Legs */}
      <mesh position={[-0.4, 0.25, 0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.05, 0.03, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.8} roughness={0.2} />
      </mesh>
      <mesh position={[0.4, 0.25, 0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.05, 0.03, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.8} roughness={0.2} />
      </mesh>
      <mesh position={[-0.4, 0.25, -0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.05, 0.03, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.8} roughness={0.2} />
      </mesh>
      <mesh position={[0.4, 0.25, -0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.05, 0.03, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.8} roughness={0.2} />
      </mesh>
    </group>
  );
}

export default function Scene({ colors }: { colors: Record<string, string> }) {
  return (
    <div className="canvas-container">
      <Canvas shadows camera={{ position: [2, 2, 3], fov: 45 }}>
        <color attach="background" args={['#0f1115']} />
        
        <ambientLight intensity={0.5} />
        <spotLight position={[5, 5, 5]} angle={0.15} penumbra={1} intensity={1} castShadow />
        <directionalLight position={[-5, 5, -5]} intensity={0.5} />
        
        <Suspense fallback={null}>
          <ChairModel colors={colors} />
          <Environment preset="city" />
          <ContactShadows position={[0, -0.5, 0]} opacity={0.4} scale={10} blur={2} far={4} />
        </Suspense>
        
        <OrbitControls 
          enablePan={false} 
          minPolarAngle={Math.PI / 4} 
          maxPolarAngle={Math.PI / 2} 
          minDistance={2} 
          maxDistance={5} 
        />
      </Canvas>
    </div>
  );
}
