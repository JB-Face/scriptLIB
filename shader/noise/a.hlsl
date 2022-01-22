
float size = 2;
float3 _BorderColor = (1.0,0.0,1.0);
float2 offset = (0,0);


float rand(float2 co){
    return frac(sin(dot(co, float2 (12.9898, 78.233))) * 43758.5453);
}

float3 rand3(float3 co){
    return frac(sin(dot(co, float3 (12.9898, 78.233,13.456))) * 43758.5453);
}

float rand1(float3 co){
    return frac(sin(dot(co, float1 (12.9898))) * 43758.5453);
}


float voronoiNoise01(float2 value){
    float2 cell = floor(value);
    float2 cellPosition = cell + rand(cell);
    float2 toCell = cellPosition - value;
    float distToCell = length(toCell);
    return distToCell;
}

float voronoiNoise(float2 value){
    float2 baseCell = floor(value);

    float minDistToCell = 10;
    for(int x=-1; x<=1; x++){
        for(int y=-1; y<=1; y++){
            float2 cell = baseCell + float2(x, y);
            float2 cellPosition = cell + rand(cell);
            float2 toCell = cellPosition - value;
            float distToCell = length(toCell);
            if(distToCell < minDistToCell){
                minDistToCell = distToCell;
            }
        }
    }
    return minDistToCell;
}

float3 voronoiNoise3d(float3 value){
    float3 baseCell = floor(value);

    //first pass to find the closest cell
    float minDistToCell = 10;
    float3 toClosestCell;
    float3 closestCell;
    [unroll]
    for(int x1=-1; x1<=1; x1++){
        [unroll]
        for(int y1=-1; y1<=1; y1++){
            [unroll]
            for(int z1=-1; z1<=1; z1++){
                float3 cell = baseCell + float3(x1, y1, z1);
                float3 cellPosition = cell + rand3(cell);
                float3 toCell = cellPosition - value;
                float distToCell = length(toCell);
                if(distToCell < minDistToCell){
                    minDistToCell = distToCell;
                    closestCell = cell;
                    toClosestCell = toCell;
                }
            }
        }
    }

    //second pass to find the distance to the closest edge
    float minEdgeDistance = 10;
    [unroll]
    for(int x2=-1; x2<=1; x2++){
        [unroll]
        for(int y2=-1; y2<=1; y2++){
            [unroll]
            for(int z2=-1; z2<=1; z2++){
                float3 cell = baseCell + float3(x2, y2, z2);
                float3 cellPosition = cell + rand3(cell);
                float3 toCell = cellPosition - value;

                float3 diffToClosestCell = abs(closestCell - cell);
                bool isClosestCell = diffToClosestCell.x + diffToClosestCell.y + diffToClosestCell.z < 0.1;
                if(!isClosestCell){
                    float3 toCenter = (toClosestCell + toCell) * 0.5;
                    float3 cellDifference = normalize(toCell - toClosestCell);
                    float edgeDistance = dot(toCenter, cellDifference);
                    minEdgeDistance = min(minEdgeDistance, edgeDistance);
                }
            }
        }
    }

    float random = rand1(closestCell);
    return float3(minDistToCell, random, minEdgeDistance);
}

// float4 main(in float2 uv:TEXCOORD0):SV_TARGET
// {

//     float2 value = uv * size;
//     float noise = voronoiNoise(value);
//     return (uv,1.0,1.0);
// }

float4 main(in float2 uv:TEXCOORD0):SV_TARGET
{

    float2 value = (uv+offset)*10 * size;
    float noise = voronoiNoise(value);
    return (noise);
}


// float4 main(in float4 worldPos:SV_POSITION):SV_TARGET
// {
//     float4 WP = worldPos/1000000;
//     float4 value = WP.xyzw/size;
//     float3 noise = voronoiNoise3d(value.xyz);

//     float3 cellColor = rand3(noise.y); 
//     float valueChange = fwidth(value.z) * 0.5;
//     float isBorder = 1 - smoothstep(0.05 - valueChange, 0.05 + valueChange, noise.z);
//     float3 color = lerp(cellColor, _BorderColor, isBorder);
//     return (noise,1.0f);
// }