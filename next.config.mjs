import withPWAInit from "@ducanh2912/next-pwa";

  const withPWA = withPWAInit({
    est: "public",
    cacheOnFrontEndNav: true ,
    aggressiveFrontEndNavCaching: true , 
    reloadOnOnline: true ,
    swcMinify: true,
    disable: false,
    workboxOptions :{
        disableDevLogs:true,
    }
  });


/** @type {import('next').NextConfig} */


const nextConfig = withPWA({});

export default nextConfig

