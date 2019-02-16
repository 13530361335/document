---
title: JDK源码阅读顺序
date: 2018-12-31 00:04:45
tags: [源码]
categories: Java
---

很多java开发的小伙伴都会阅读jdk源码，然而确不知道应该从哪读起。
以下为小编整理的通常所需阅读的源码范围,标题为包名，后面序号为优先级1-4，优先级递减。

**1、java.lang**
``` bash
1) Object 1
2) String 1
3) AbstractStringBuilder 1
4) StringBuffer 1
5) StringBuilder 1
6) Boolean 2
7) Byte 2
8) Double 2
9) Float 2
10) Integer 2
11) Long 2
12) Short 2
13) Thread 2
14) ThreadLocal 2
15) Enum 3
16) Throwable 3
17) Error 3
18) Exception 3
19) Class 4
20) ClassLoader 4
21) Compiler 4
22) System 4
23) Package 4
24) Void 4
```

**2、java.util**
``` bash
1) AbstractList 1
2) AbstractMap 1
3) AbstractSet 1
4) ArrayList 1
5) LinkedList 1
6) HashMap 1
7) Hashtable 1
8) HashSet 1
9) LinkedHashMap 1
10) LinkedHashSet 1
11) TreeMap 1
12) TreeSet 1
13) Vector 2
14) Queue 2
15) Stack 2
16) SortedMap 2
17) SortedSet 2
18) Collections 3
19) Arrays 3
20) Comparator 3
21) Iterator 3
22) Base64 4
23) Date 4
24) EventListener 4
25) Random 4
26) SubList 4
27) Timer 4
28) UUID 4
29) WeakHashMap 4
```

**3、java.util.concurrent**
``` bash
1) ConcurrentHashMap 1
2) Executor 2
3) AbstractExecutorService 2
4) ExecutorService 2
5) ThreadPoolExecutor 2
6) BlockingQueue 2
7）AbstractQueuedSynchronizer 2
8）CountDownLatch 2
9) FutureTask 2
10）Semaphore 2
11）CyclicBarrier 2
13）CopyOnWriteArrayList 3
14）SynchronousQueue 3
15）BlockingDeque 3
16) Callable 4
```

**4、java.util.concurrent.atomic**
``` bash
1) AtomicBoolean 2
2) AtomicInteger 2
3) AtomicLong 2
4) AtomicReference 3
```

**5、java.lang.reflect**
``` bash
1) Field 2
2) Method 2
```

**6、java.lang.annotation**
``` bash
1) Annotation 3
2) Target 3
3) Inherited 3
4) Retention 3
5) Documented 4
6) ElementType 4
7) Native 4
8) Repeatable 4
```

**7、java.util.concurrent.locks**
``` bash
1) Lock 2
2) Condition 2
3) ReentrantLock 2
4) ReentrantReadWriteLock 2
```

**8、java.io**
``` bash
1) File 3
2) InputStream   3
3) OutputStream  3
4) Reader  4
5) Writer  4
```

**9、java.nio**
``` bash
1) Buffer 3
2) ByteBuffer 4
3) CharBuffer 4
4) DoubleBuffer 4
5) FloatBuffer 4
6) IntBuffer 4
7) LongBuffer 4
8) ShortBuffer 4
```

**10、java.sql**
``` bash
1) Connection 3
2) Driver 3
3) DriverManager 3
4) JDBCType 3
5) ResultSet 4
6) Statement 4
```

**11、java.net**
``` bash
1) Socket 3
2) ServerSocket 3
3) URI 4
4) URL 4
5) URLEncoder 4
```

###阅读笔记简版
1、Object 
1) wait(), notify(), notifyAll(), wait(timeout) 
2) hashCode(), equals() 
3) clone()

2、String 
1) char[] value 
2) int hash 
3) equals(), startWith(), endWith(), replace

3、AbstractStringBuilder 
1) char[] value 
2) int count 
3) 扩容：翻倍，不够取所需最小

4、StringBuffer 
1) 继承AbstractStringBuilder 
2) synchronized方法保证线程安全 
3) char[] toStringCache

5、StringBuilder 继承AbstractStringBuilder

6、ArrayList 
1) Object[] elementData 
2) int size 
3) 默认大小10 
4) 扩容：翻倍，不够取所需最小

7、LinkedList 
1) Node {E item, Node prev, Node next} 
2) int size 
3) Node first 
4) Node last 
5) linkFirst(), linkLast(), linkBefore(), unLinkFirst(), unLinkLast(), unLink(), indexOf()

8、HashMap 
1) Node{int hash, K key, V value, Node next} 
2) 默认容量16，负载因子0.75f 
3) int size, modCount, threshold, float loadFactor 
4) Node[] table 
5) Set entrySet 
6) put():根据key算hash，根据容量和hash算index，table[index]没有直接添加到数组中，table[index]有，若index位置同一个key则更新，否则遍历next是否有，有则更新，无则新增，最后根据thread与size判断是否扩容。注：扩容时容量翻倍，重新算hash复制到新数组 
7）get()类似 
注：先比较hash，若相等在比较equals

9、Hashtable 
1) 结构实现与HashMap基本一致 
2)通过synchronized方法保证线程安全

10、HashSet：委托给HashMap，其Value是同一个默认对象

11、LinkedHashMap继承HashMap 
1) Entry{HashMap.Node, Entry before, after} 
2) Entry head, tail 
3) 重写newNode()添加节点时，除像HashMap中添加外，保存before、after信息

12、LinkedHashSet继承HashSet：不知道如何实现的顺序？

13、AbstractMap维护EntrySet，AbstractSet维护Iterator，AbstractList维护Iterator

14、ConcurrentHashMap 
1) JDK1.7及以前： 
a、Segment[] ,HashEntry[] , HashEntry{hash, k, v, next} 
b、根据key算hash，根据hash和Segment的大小算位置，每个segment拥有一个自己的HashEntry[] 
c、get()：不加锁，volatile类型 
d、put(): 对相应segment加锁 
e、size()：各HashEntry[] 之和，先不加锁算两遍，若一致则返回，若不一致则加锁重新计算 
2）JDK1.8 
a、Node{hash, key, value, next} 
b、Node[] table 
c、大多数操作类似于HashMap，不同CAS方式设置，根据key算hash，在根据hash和容量算index，对table[index]加锁，从而达到更大的并发量 
d、get(): 同HashMap 
e、put(): 对table[index]加锁

15、TreeMap 
1）红黑树，即自平衡二叉查找树，时间复杂度O(logn) 
2）Entry{K k, V v, Entry parent, left, right, boolean color} 
3）Entry root，int size， int modeCount

16、TreeSet：委托TreeMap实现
--------------------- 
作者：猴子哥哥1024 
来源：CSDN 
原文：https://blog.csdn.net/qq_21033663/article/details/79571506 
版权声明：本文为博主原创文章，转载请附上博文链接！