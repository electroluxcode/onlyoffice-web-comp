'use client'

<<<<<<< HEAD
import StudioLayout from '@/components/StudioLayout';
=======
import StudioLayout from '@/components/studio-layout';
>>>>>>> refactor/v9

export default function ClientLayoutWrapper({
  children,
}: {
  children: React.ReactNode;
}) {
  return <StudioLayout>{children}</StudioLayout>;
}
