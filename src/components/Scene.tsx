'use client';

import { Canvas } from '@react-three/fiber';
import { OrbitControls, Environment, ContactShadows, PresentationControls } from '@react-three/drei';
import { Suspense } from 'react';

function ChairModel({ colors }: { colors: Record<string, string> }) {
  return (
    <group position={[0, -0.6, 0]}>
      {/* Seat */}
      <mesh position={[0, 0.5, 0]} castShadow receiveShadow>
        <boxGeometry args={[1, 0.1, 1]} />
        <meshStandardMaterial color={colors.seat || '#e0e0e0'} roughness={0.6} />
      </mesh>
      
      {/* Backrest */}
      <mesh position={[0, 1.1, -0.45]} castShadow receiveShadow>
        <boxGeometry args={[1, 0.8, 0.1]} />
        <meshStandardMaterial color={colors.backrest || '#333333'} roughness={0.7} />
      </mesh>
      
      {/* Legs */}
      <mesh position={[-0.4, 0.25, 0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.04, 0.02, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.9} roughness={0.1} />
      </mesh>
      <mesh position={[0.4, 0.25, 0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.04, 0.02, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.9} roughness={0.1} />
      </mesh>
      <mesh position={[-0.4, 0.25, -0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.04, 0.02, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.9} roughness={0.1} />
      </mesh>
      <mesh position={[0.4, 0.25, -0.4]} castShadow receiveShadow>
        <cylinderGeometry args={[0.04, 0.02, 0.5]} />
        <meshStandardMaterial color={colors.legs || '#111111'} metalness={0.9} roughness={0.1} />
      </mesh>
    </group>
  );
}

export default function Scene({ colors }: { colors: Record<string, string> }) {
  return (
    <div className="canvas-container">
      <Canvas shadows camera={{ position: [3, 2.5, 4], fov: 40 }}>
        <color attach="background" args={['#0a0a0a']} />
        
        {/* Premium Lighting Setup */}
        <ambientLight intensity={0.4} />
        <spotLight position={[5, 5, 5]} angle={0.2} penumbra={1} intensity={1.5} castShadow shadow-bias={-0.0001} />
        <spotLight position={[-5, 5, -5]} angle={0.2} penumbra={1} intensity={0.8} />
        
        <Suspense fallback={null}>
          <PresentationControls 
            global 
            rotation={[0, 0.3, 0]} 
            polar={[-Math.PI / 4, Math.PI / 4]} 
            azimuth={[-Math.PI / 2, Math.PI / 2]} 
            config={{ mass: 2, tension: 500 }} 
            snap={{ mass: 4, tension: 1500 }}
          >
            <ChairModel colors={colors} />
          </PresentationControls>
          
          {/* Photorealistic Environment Reflections */}
          <Environment preset="studio" />
          
          {/* Realistic Contact Shadows */}
          <ContactShadows position={[0, -0.6, 0]} opacity={0.6} scale={10} blur={2} far={4} color="#000000" />
        </Suspense>
        
        {/* Fallback controls if PresentationControls doesn't suffice, but PresentationControls is smoother for product viewers */}
        {/* <OrbitControls enablePan={false} minPolarAngle={0} maxPolarAngle={Math.PI / 2 + 0.1} minDistance={2} maxDistance={6} makeDefault /> */}
      </Canvas>
    </div>
  );
}
