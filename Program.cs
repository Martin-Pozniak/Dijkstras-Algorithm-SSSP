/************************************************************************************************************************************
* Author: Martin Pozniak
* Date: 12/6/17
* Description: Parses a formatted excel spreadsheet containing a source and destination node followed by a weight joining the nodes.
*              Implements Dijkstra's Algorithm to complete the SSSP Task.
************************************************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

/**********************************************
* NameSpace: SSSP
**********************************************/
namespace SSSP
{
    /**********************************************
    * Class: Vertex - Contains definitions and data for each vertex representing a city
    **********************************************/
    class Vertex : IEquatable<Vertex>
    {
        /**********************************************
        * Member Variables
        **********************************************/
        private string name;
        private Vertex previous;
        private List<Edge> edges= new List<Edge>();
        int distanceFromSource = int.MaxValue;

        public Vertex()
        {
        }
        public Vertex(string n)
        {
            name = n;
        }
        public void addName(string name)
        {
            this.name = name;
        }
        public void addEdge(Edge e)
        {
            edges.Add(e);
        }
        public string getName()
        {
            return name;
        }
        public List<Edge> getEdges()
        {
            return edges;
        }
        public void addEdge(Vertex s, Vertex d, int w)
        {
            Edge newEdge = new Edge(s, d, w);
            edges.Add(newEdge);
        }
        public void setDistanceFromSource(int d)
        {
            distanceFromSource = d;
        }
        public void setPrevious(Vertex prev)
        {
            previous = prev;
        }
        public Vertex getPrevious()
        {
            return previous;
        }
        public bool hasEdgeTo(Vertex dest)
        {
            if (edges.Any(p => p.getDest() == dest)) return true;
            else return false;
        }
        public bool Equals(Vertex other)
        {
            if (name.CompareTo(other.name) == 0) return true;

            return false;
               
        }

        public int getDistanceFromSource()
        {
            return distanceFromSource;
        }
    }

    /**********************************************
    * Class: Edge - Contains definitions and data for each edge joining each vertex
    **********************************************/
    class Edge
    {
        /**********************************************
        * Member Variables
        **********************************************/
        private Vertex source;
        private Vertex dest;
        private int weight;

        public Edge(Vertex source, Vertex dest, int weight)
        {
            this.source = source;
            this.dest = dest;
            this.weight = weight;
        }       
        public Vertex getSource()
        {
            return source;
        }
        public Vertex getDest()
        {
            return dest;
        }
        public int getWeight()
        {
            return weight;
        }
    }

    /**********************************************
    * Class: Graph - an Adjacency list structured graph containing information of all edges and vertices and therefore creating a graph
    *        pg 361, performs all operations listed on page 360
    **********************************************/
    class Graph 
    {
        /**********************************************
        * Member Variables
        **********************************************/
        private List<Vertex> VertexList=null;
        private List<List<Edge>> adjacencyList = new List<List<Edge>>();
        private int noVertices = 0;
        private int noEdges = 0;

        public Graph(List<Vertex> v, List<Edge> e)
        {
            noVertices = v.Count;
            noEdges = e.Count;
            VertexList = v;
            foreach (Vertex vert in VertexList)
            {
                adjacencyList.Add(vert.getEdges()); //adjacency list contains a list of List of edges for each node v
            }
        }       
        public int getNoVertices()
        {
            return noVertices;
        }
        public int getNoEdges()
        {
            return noEdges;
        }
        public List<Vertex> getAllVertices()
        {
            return VertexList;
        }
        public List<List<Edge>> getAllEdges()
        {
            return adjacencyList;
        }      
        public int getVertexDegree(string name)
        {
            int degree=0;
            Vertex v = null;
            v = VertexList[VertexList.FindIndex(f => f.getName().CompareTo(name) == 0)]; ;
            foreach(Edge e in v.getEdges())
            {
                degree++;
            }
            return degree * 2;
        }
        public List<Edge> getIncidentEdgesOn(string name)
        {
            Vertex v = null;
            v = VertexList[VertexList.FindIndex(f => f.getName().CompareTo(name) == 0)]; ;
            return v.getEdges();
        }
        public List<Vertex> getNeighborVerticesOf(string name)
        {
            Vertex v = null;
            List<Vertex> neighbors = new List<Vertex>();
            v = VertexList[VertexList.FindIndex(f => f.getName().CompareTo(name) == 0)];
            foreach(Edge e in v.getEdges())
            {
                neighbors.Add(e.getDest());
            }
            return neighbors;
        }
        public bool areAdjacent(string source, string dest)
        {
            Vertex v1 = VertexList[VertexList.FindIndex(f => f.getName().CompareTo(source) == 0)];
            Vertex v2 = VertexList[VertexList.FindIndex(f => f.getName().CompareTo(dest) == 0)];
            if (v1.hasEdgeTo(v2))
            {
                return true;
            }
             return false;
        }
    }

    /**********************************************
    * Class: Program - Main Program Class
    **********************************************/
    class Program
    {
        /**********************************************
        * Member Variables
        **********************************************/
        private static List<Vertex> VertexList = new List<Vertex>();
        private static List<Edge> Edges = new List<Edge>();
        private static int totalDistance = 0;
        /// <summary>
        /// Set the Following Constant OP to one of the following values to have the main program run a different sequence of functions
        ///     
        ///     IMPORTANT: Change the path variable to the path of the data file before running the program or the program will fail to find the data-file and quit. The path is specified on line 466
        /// 
        ///     1) Runs Dijkstra's to satisy pt 4.1-4.3 of Task IV in the assignment. Finds 7 individual routes
        ///     2) Runs Dijkstra's to find shortest path from 'SOURCE_NODE' to every other node in the graph (SSSP)
        ///     3) Prints out all details of the graph based on the graph functions listed on pg 360 that we didnt use in other parts of the code
        /// </summary>

        private static string PATH_TO_DATA ="F:\\CS242 Final Project\\DATA-FINAL-REMADE.xlsx";
        private static int OP= 1;
        private static string SOURCE_NODE = "Chicago";

        /**********************************************
        * Main Function
        **********************************************/
        static void Main(string[] args)
        {

            List<Vertex> path = new List<Vertex>();
            parseExcelFile();
            Graph G = new Graph(VertexList, Edges);

            if ( OP == 1 )
            {
                path = RunDijkstrasOnRoutes( G );
                Console.WriteLine();
                Console.WriteLine("Total Distance Over The 7 Routes: " + totalDistance);
            }
            else if ( OP == 2 )
            {
                path = getShortestToAllNodes(G, getVertexInList( SOURCE_NODE ) );
                foreach(Vertex v in path)
                {
                    Console.WriteLine("--------------------------------------------------------------------------");
                    Console.WriteLine("Shortest Distance To "+v.getName()+" From "+getVertexInList(SOURCE_NODE).getName());
                    Console.WriteLine("--------------------------------------------------------------------------");
                    printPathAndData(backTrackForPathFrom(v));
                }
            }
            else if ( OP == 3 )
            {
                printAllDataAbout(G);
            }
                     
            Console.WriteLine("\n-Done-");
            Console.ReadLine();
        }

        private static void printAllDataAbout(Graph G)
        {
            Console.WriteLine("---------PRINTING RELEVENT GRAPH INFO-----------");
            Console.WriteLine("No. Vertices: " + G.getNoVertices());
            Console.WriteLine("No. Edges: "+G.getNoEdges());
            Console.WriteLine("Degree of "+SOURCE_NODE+": "+G.getVertexDegree(SOURCE_NODE));
            Console.WriteLine("Are Chicago and Milwaukee Adjacent? " + G.areAdjacent("Chicago", "Milwaukee"));
            Console.WriteLine("Are Chicago and Grand Forks Adjacent? " + G.areAdjacent("Chicago", "Grand Forks"));
        }

        private static List<Vertex> RunDijkstrasOnRoutes(Graph G)
        {
            List<Vertex> path;
            Console.WriteLine("------------------------Route 1-------------------------");
            path = getShortestPathFromSourceToDest(G, getVertexInList("Grand Forks"), getVertexInList("Seattle"));
            printPathAndData(path);
            Console.WriteLine();
            Console.WriteLine("------------------------Route 2-------------------------");
            path = getShortestPathFromSourceToDest(G, getVertexInList("Seattle"), getVertexInList("Los Angeles"));
            printPathAndData(path);
            Console.WriteLine();
            Console.WriteLine("------------------------Route 3-------------------------");
            path = getShortestPathFromSourceToDest(G, getVertexInList("Los Angeles"), getVertexInList("Dallas"));
            printPathAndData(path);
            Console.WriteLine();
            Console.WriteLine("------------------------Route 4-------------------------");
            path = getShortestPathFromSourceToDest(G, getVertexInList("Dallas"), getVertexInList("Miami"));
            printPathAndData(path);
            Console.WriteLine();
            Console.WriteLine("------------------------Route 5-------------------------");
            path = getShortestPathFromSourceToDest(G, getVertexInList("Miami"), getVertexInList("Boston"));
            printPathAndData(path);
            Console.WriteLine();
            Console.WriteLine("------------------------Route 6-------------------------");
            path = getShortestPathFromSourceToDest(G, getVertexInList("Boston"), getVertexInList("Chicago"));
            printPathAndData(path);
            Console.WriteLine();
            Console.WriteLine("------------------------Route 7-------------------------");
            path = getShortestPathFromSourceToDest(G, getVertexInList("Chicago"), getVertexInList("Grand Forks"));
            printPathAndData(path);
            Console.WriteLine();
            return path;
        }

        private static void printPathAndData(List<Vertex> pathFromSourceToDest)
        {
            totalDistance += pathFromSourceToDest[0].getDistanceFromSource();
            pathFromSourceToDest.Reverse(); //This is why we love c# since we were backtracking our path is from dest to source...so flip it
            foreach (Vertex v in pathFromSourceToDest)
            {
                Console.Write("-" + v.getName() + "-"+v.getDistanceFromSource()+"-");               
            }
            Console.WriteLine("\n----------------Route SubTotal:" + pathFromSourceToDest[pathFromSourceToDest.Count - 1].getDistanceFromSource()+"------------------------");
        }

        public static Vertex getVertexInList(string name)
        {
            Vertex CopyOfOriginal = new Vertex();
            CopyOfOriginal = VertexList[VertexList.FindIndex(f => f.getName().CompareTo(name) == 0)];
            return CopyOfOriginal;
        }

        public static List<Vertex> getShortestPathFromSourceToDest(Graph G, Vertex source,Vertex dest) //The difference here is that once we remove 'dest' from the fringe and add it to the cloud we know we found the shortest path to it. then we backtrack to provide the path
        {
            List<Vertex> cloud = new List<Vertex>();
            List<Vertex> fringe = new List<Vertex>();//we define a priority queue 'fringe' that initially contains all the nodes outside of the cloud  

            //Set every other node in the graph to have inf distance... this allows us to reuse the algorithm
            foreach (Vertex v in G.getAllVertices())
            {
                v.setDistanceFromSource(int.MaxValue);
            }

            fringe = G.getAllVertices();
            source.setPrevious(null);
            source.setDistanceFromSource(0);

            while (fringe.Count > 0)
            {

                fringe = fringe.OrderBy(f => f.getDistanceFromSource()).ToList(); 

                foreach (Edge w in G.getIncidentEdgesOn(fringe[0].getName()))
                {
                    relaxEdge(cloud, fringe, fringe[0], w);
                }
                cloud.Add(fringe[0]);                

                if (fringe[0].getName().CompareTo(dest.getName()) == 0)//We found the shortest path to dest node
                {
                    return backTrackForPathFrom(fringe[0]);
                }
                fringe.RemoveAt(0);
            }

            //Console.WriteLine("Fringe Empty So All Shorest Paths From " + source.getName() + " to every other node has been found and are stored in the cloud");

            return cloud;
        }

        private static void relaxEdge(List<Vertex> cloud, List<Vertex> fringe, Vertex m, Edge w)
        {
            
            if (w.getDest().getDistanceFromSource() == int.MaxValue)
            {
                w.getDest().setDistanceFromSource(m.getDistanceFromSource() + w.getWeight());
                //fringe.Add(w.getDest());
                w.getDest().setPrevious(m);
            }
            else
            {//Edge Relaxation
                int v1 = w.getDest().getDistanceFromSource() + w.getWeight();
                int v2 = m.getDistanceFromSource();
                m.setDistanceFromSource(Min(v1, v2));
                //we need to set the previous of e.getDest = whichever node happened to provide the lesser distance from source
                if (v1 > v2)//if the distance that the node prev had from source is less than the new route don't change it
                { } //NO-OP                          
                else//otherwise set the new previous of the node to the intermediate node
                {
                    m.setPrevious(w.getDest());
                }
            }
            
        }

        public static List<Vertex> getShortestToAllNodes(Graph G, Vertex source) //Input, A weighted undirected graph G, and a source Vertex OutPut: a list of all nodes with their shortes paths to source
        {
            List<Vertex> cloud = new List<Vertex>();
            List<Vertex> fringe = new List<Vertex>(); //we define a priority queue 'fringe' that initially contains all the nodes outside of the cloud

            //Set every other node in the graph to have inf distance... this allows us to reuse the algorithm
            foreach (Vertex v in G.getAllVertices())
            {
                v.setDistanceFromSource(int.MaxValue);
            }
            fringe = G.getAllVertices();
            source.setPrevious(null);
            source.setDistanceFromSource(0);
        
            while (fringe.Count > 0)
            {
                //order a priority queue 'fringe' by distance from source of the fringe and remove the closest node to source which will be fringe[0]                
                fringe = fringe.OrderBy(f => f.getDistanceFromSource()).ToList();
                                           
                foreach (Edge w in G.getIncidentEdgesOn(fringe[0].getName()))
                {
                    relaxEdge(cloud, fringe, fringe[0], w);
                }
                
                cloud.Add(fringe[0]);
                fringe.RemoveAt(0);
            }

            Console.WriteLine("Shorest Paths From " + source.getName() + " to every other node has been found and are stored in the cloud");

            return cloud;
        }

        public static List<Vertex> backTrackForPathFrom(Vertex v) //pass in the destination and we will return a list of nodes that backtrack to the source
        {
            List<Vertex> pathToSource = new List<Vertex>();
            Vertex current = v;
            while(current!=null)
            {
                pathToSource.Add(current);
                current = current.getPrevious();
            }
            return pathToSource; //path to source is backwards
        }

        private static int Min(int v1, int v2)
        {
            if (v1 < v2) return v1;
            else return v2;
        }

        public static void parseExcelFile()
        {         
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            string colData;
            string sourceName = "";
            string destName = "";
            int weight = 0;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;
            
            //List<Vertex> nodesToAdd = new List<Vertex>();
            bool sourceInList = false;
            bool destInList = false;
            int sourceIndex = -1;
            int destIndex = -1;


            xlApp = new Excel.Application();

            /****************************************************************************************
             * IMPORTANT CAHNGE THIS PATH BELOW TO THE PATH OF THE DATA FILE ON YOUR FILE SYSTEM
            ****************************************************************************************/
            xlWorkBook = xlApp.Workbooks.Open(@PATH_TO_DATA, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);           
            //************************************************************************************
            //************************************************************************************

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;          

            for (rCnt = 1; rCnt <= rw; rCnt++)//for each row
            {
                //Console.WriteLine();              
                sourceInList = false;
                destInList = false;

                Vertex newVertex=new Vertex();
                Vertex newDestVertex = new Vertex();                

                for (cCnt = 1; cCnt <= cl; cCnt++) //get column data
                {
                    colData = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Text;
                    if (cCnt == 1)//Source Name  
                    {
                        //Console.Write(colData);
                        sourceName = colData;
                    }
                    else if (cCnt == 2) //Dest Name
                    {
                        //Console.Write(colData);
                        destName = colData;
                    }
                    else if (cCnt == 3) //weight
                    {
                        //Console.Write(colData);
                        Int32.TryParse(colData, out weight);
                    }
                }
                Edges.Add(new Edge(newVertex, newDestVertex, weight));
                newVertex.addName(sourceName);
                newDestVertex.addName(destName);

                //check if source and dest are in the list and if so assign them the index in the list for easy access
                if(VertexList.Contains(newVertex))
                {
                    sourceInList = true;
                    sourceIndex = VertexList.FindIndex(f => f.getName().CompareTo(sourceName)==0);
                }
                if (VertexList.Contains(newDestVertex))
                {
                    destInList = true;
                    destIndex = VertexList.FindIndex(f => f.getName().CompareTo(destName) == 0);
                }
                
                //now we have the source,dest,and weight. We can now create our add to our vertices and edge lists
                //VertexList.Any(p => p.getName().CompareTo(sourceName)==0)
                if ( sourceInList )//if the VertexListAlready Contains This City
                {
                    //Console.WriteLine("List Already Contains " + sourceName);
                    if( destInList )//If both source and dest are in the list already
                    {
                        //Console.WriteLine("List Also Already Contains " + destName);  
                        if(VertexList[sourceIndex].hasEdgeTo(VertexList[destIndex]))
                        {
                            //Console.WriteLine(sourceName + " already has edge to " + destName);                           
                        }    
                        else //source doesn't have edge to dest
                        {
                            //Console.WriteLine("Adding Edge From" + sourceName + " to" + destName);
                            VertexList[sourceIndex].addEdge(new Edge(VertexList[sourceIndex], VertexList[destIndex], weight));
                            //Edges.Add(new Edge(newDestVertex, newVertex, weight));
                        }
                        if (VertexList[destIndex].hasEdgeTo(VertexList[sourceIndex]))
                        {
                            //Console.WriteLine(destName + " also already has edge to " + sourceName + " so NO-OP");
                        }
                        else
                        {
                            //Console.WriteLine("Adding Edge from " + destName + " to " + sourceName);
                            VertexList[destIndex].addEdge(new Edge(VertexList[destIndex], VertexList[sourceIndex], weight));
                            //Edges.Add(new Edge(newDestVertex, newVertex, weight));
                        }
                    }
                    else//if the destination is not in the list but source is
                    {
                        //Console.WriteLine("Adding " + destName);
                        VertexList.Add(newDestVertex);
                        //Console.WriteLine("Adding Edge From " + destName + " to " + sourceName );                
                        newDestVertex.addEdge(new Edge(newDestVertex, VertexList[sourceIndex], weight));
                        //Edges.Add(new Edge(newDestVertex, newVertex, weight));
                        if (VertexList[sourceIndex].hasEdgeTo(VertexList[VertexList.FindIndex(f => f.getName().CompareTo(destName) == 0)])) //special condition where we just added the destination node and need to check something about it
                        {
                            //Console.WriteLine(sourceName + " already has edge to " + destName);
                        }
                        else
                        {
                            //Console.WriteLine("Adding Edge from "+sourceName + " to " + destName);
                            VertexList[sourceIndex].addEdge(new Edge(VertexList[sourceIndex], VertexList[VertexList.FindIndex(f => f.getName().CompareTo(destName) == 0)], weight));//special condition where we just added the destination node and need to check something about it
                            //Edges.Add(new Edge(newVertex, newDestVertex, weight));
                        }
                        
                    }
                }
                else //if the source node we are reading is not already in the list
                {
                    //Console.WriteLine("Adding " + sourceName);
                    //we need to add it to the list                   
                    VertexList.Add(newVertex);                
                    if (destInList) //is the destination already in the list?
                    {
                        newVertex.addEdge(new Edge(newVertex, VertexList[destIndex], weight));
                        //does it have an edge to source already?
                        if (VertexList[destIndex].hasEdgeTo(newVertex))
                        {
                            //Console.WriteLine(destName+" Already Has an Edge to "+sourceName);
                            //NO-OP
                        }
                        else//no it doesnt
                        {
                            //Console.WriteLine("Adding new Edge To " + sourceName +" From " + destName);
                            VertexList[destIndex].addEdge(new Edge(VertexList[destIndex], newVertex, weight));
                            //Edges.Add(new Edge(newDestVertex, newVertex, weight));
                        }
                        
                    }
                    else //neither source or dest in list so add the dest as well
                    {
                        //Console.WriteLine("Adding "+destName);
                        //Console.WriteLine("Adding Edge from " + destName + " to " + sourceName);
                        VertexList.Add(newDestVertex);
                        newDestVertex.addEdge(new Edge(newDestVertex, newVertex, weight));
                        //Edges.Add(new Edge(newDestVertex, newVertex, weight));
                        //Console.WriteLine("Adding Edge From " + sourceName + " to " + destName);
                        newVertex.addEdge(new Edge(newVertex, newDestVertex, weight));
                        //Edges.Add(new Edge(newVertex, newDestVertex, weight));
                    }

                }

               

            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
